"""CDF_helper - Flask web app (报关清单生成器).

Run with:  python main.py        (opens browser at http://127.0.0.1:5000)
        or  flask --app webapp run
"""

import datetime
import shutil
import time
import uuid
from pathlib import Path

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for

from cdf_helper.generator import generate, sanitize_filename
from cdf_helper.parser import parse_sources, detect_vessel

REPORT_LABEL = "报关清单"
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
GENERATED_DIR = BASE_DIR / "generated"
UPLOAD_DIR.mkdir(exist_ok=True)
GENERATED_DIR.mkdir(exist_ok=True)

RETENTION_DAYS = 7


def _cleanup_old_files(directory, max_age_days=RETENTION_DAYS):
    """删除超过 max_age_days 天的旧文件，避免 uploads/generated 目录无限累积。"""
    cutoff = time.time() - max_age_days * 86400
    removed = 0
    for f in directory.glob("*"):
        if f.is_file() and f.stat().st_mtime < cutoff:
            try:
                f.unlink()
                removed += 1
            except OSError:
                pass
    return removed


_cleanup_old_files(UPLOAD_DIR)
_cleanup_old_files(GENERATED_DIR)

app = Flask(__name__)
app.config["SECRET_KEY"] = "cdf-helper"
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB upload limit


def _spreadsheets_in(directory):
    return sorted(directory.glob("*.xls*"))


def _server_source_files():
    """Source candidates: spreadsheets in the project root, excluding the template & outputs."""
    files = _spreadsheets_in(BASE_DIR)
    template = _server_template_file()
    return [f for f in files if f != template and f.parent == BASE_DIR]


def _server_template_file():
    for f in _spreadsheets_in(BASE_DIR):
        if f.parent == BASE_DIR and "报关清单" in f.name:
            return f
    return None


def _save_upload(file_storage, subdir):
    """Save an uploaded file into UPLOAD_DIR (mirroring the original filename)."""
    name = sanitize_filename(file_storage.filename or "upload")
    dest = UPLOAD_DIR / name
    if dest.exists():
        dest = UPLOAD_DIR / f"{dest.stem}-{uuid.uuid4().hex[:6]}{dest.suffix}"
    file_storage.save(dest)
    return dest


@app.template_filter("basename")
def _basename(path):
    return Path(str(path)).name


@app.route("/")
def index():
    return render_template(
        "index.html",
        template_candidates=[f for f in _spreadsheets_in(BASE_DIR) if f.suffix == ".xlsx"],
        server_template=_server_template_file(),
        source_candidates=_server_source_files(),
        today=datetime.date.today().isoformat(),
    )


@app.route("/generate", methods=["POST"])
def do_generate():
    # --- template -------------------------------------------------------
    template = None
    uploaded_template = request.files.get("template_upload")
    if uploaded_template and uploaded_template.filename:
        template = _save_upload(uploaded_template, UPLOAD_DIR)
    else:
        template_path = request.form.get("template_path", "").strip()
        if template_path:
            cand = (BASE_DIR / template_path).resolve()
            if cand.is_file() and cand.suffix.lower() in (".xls", ".xlsx"):
                template = cand
    if template is None or not template.is_file():
        flash("请选择或上传报关清单模板文件。", "error")
        return redirect(url_for("index"))

    # --- sources --------------------------------------------------------
    sources = []
    for f in request.files.getlist("sources_upload"):
        if f and f.filename:
            sources.append(_save_upload(f, UPLOAD_DIR))
    for sp in request.form.getlist("source_paths"):
        if sp.strip():
            cand = (BASE_DIR / sp.strip()).resolve()
            if cand.is_file() and cand.suffix.lower() in (".xls", ".xlsx"):
                sources.append(cand)
    if not sources:
        flash("请至少选择一个或上传一个备件来源文件。", "error")
        return redirect(url_for("index"))

    # --- parse ----------------------------------------------------------
    warnings = []

    def warn(msg):
        warnings.append(msg)

    try:
        parts = parse_sources(sources, warn=warn)
    except ValueError as e:
        flash(str(e), "error")
        return redirect(url_for("index"))

    vessel = request.form.get("vessel", "").strip()
    if not vessel:
        vessel = detect_vessel(sources) or ""
    port = request.form.get("port", "").strip()
    date = request.form.get("date", "").strip() or datetime.date.today().isoformat()
    include_spec = request.form.get("include_spec") == "on"

    output_name = f"{sanitize_filename(vessel)}-{sanitize_filename(port)}-{REPORT_LABEL}-{sanitize_filename(date)}.xlsx"
    out_path = GENERATED_DIR / output_name
    if out_path.exists():
        out_path = GENERATED_DIR / f"{out_path.stem}-{uuid.uuid4().hex[:6]}{out_path.suffix}"

    try:
        generate(
            template_path=template,
            parts=parts,
            vessel_name=vessel,
            output_dir=GENERATED_DIR,
            output_name=out_path.name,
            include_spec=include_spec,
        )
    except Exception as e:
        flash(f"生成失败：{e}", "error")
        return redirect(url_for("index"))

    return render_template(
        "result.html",
        file_name=out_path.name,
        vessel=vessel,
        port=port,
        date=date,
        item_count=len(parts),
        warnings=warnings,
    )


@app.route("/download/<path:file_name>")
def download(file_name):
    target = (GENERATED_DIR / file_name).resolve()
    if not target.is_file() or target.parent != GENERATED_DIR:
        abort(404)
    return send_file(target, as_attachment=True, download_name=target.name)


if __name__ == "__main__":
    app.run(debug=True)