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

from cdf_helper import config as app_config
from cdf_helper.ai import enrich_parts
from cdf_helper.generator import generate, sanitize_filename
from cdf_helper.parser import (
    parse_sources,
    detect_vessel,
    detect_vessel_pair,
    load_vessel_names,
    bilingual_vessel,
)
from cdf_helper.translator import translate_names

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
    """Source candidates: spreadsheets in the project root, excluding the template, outputs & the 船名 lookup."""
    files = _spreadsheets_in(BASE_DIR)
    template = _server_template_file()
    lookup = _server_vessel_lookup()
    return [f for f in files if f != template and f != lookup and f.parent == BASE_DIR]


def _server_vessel_lookup():
    """The 中英文船名 lookup workbook in the project root (name contains '船名'), or None."""
    for f in _spreadsheets_in(BASE_DIR):
        if f.parent == BASE_DIR and "船名" in f.name:
            return f
    return None


def _server_vessel_names():
    """Load (zh2en, en2zh) from the lookup workbook, falling back to empty dicts."""
    lookup = _server_vessel_lookup()
    if lookup is None:
        return {}, {}
    try:
        return load_vessel_names(lookup)
    except Exception:
        return {}, {}


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
    trans_app_id, trans_apikey = app_config.get_trans_credentials()
    return render_template(
        "index.html",
        template_candidates=[f for f in _spreadsheets_in(BASE_DIR) if f.suffix == ".xlsx"],
        server_template=_server_template_file(),
        source_candidates=_server_source_files(),
        today=datetime.date.today().isoformat(),
        api_key_prefill=app_config.get_api_key(),
        trans_app_id_prefill=trans_app_id,
        trans_apikey_prefill=trans_apikey,
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

    # --- always-on: translate English part names to Chinese --------------
    _saved_app_id, _saved_apikey = app_config.get_trans_credentials()
    trans_app_id = request.form.get("trans_app_id", "").strip() or _saved_app_id
    trans_apikey = request.form.get("trans_apikey", "").strip() or _saved_apikey
    if request.form.get("save_trans_key") == "on" and trans_app_id and trans_apikey:
        app_config.save_config({"trans_app_id": trans_app_id, "trans_apikey": trans_apikey})
    trans_stats = None
    if trans_app_id and trans_apikey:
        try:
            trans_stats = translate_names(parts, trans_app_id, trans_apikey, on_status=warn)
        except Exception as e:
            warn(f"翻译调用失败：{e}")
    else:
        warn("未配置翻译 API 凭据，跳过英文名称翻译。")

    vessel = request.form.get("vessel", "").strip()
    chinese, english = (None, None)
    if not vessel:
        chinese, english = detect_vessel_pair(sources)
        if chinese is None and english is None:
            vessel = detect_vessel(sources) or ""
    if vessel or chinese or english:
        zh2en, en2zh = _server_vessel_names()
        vessel = bilingual_vessel(vessel, zh2en, en2zh, english=english, chinese=chinese)
    port = request.form.get("port", "").strip()
    date = request.form.get("date", "").strip() or datetime.date.today().isoformat()
    include_spec = request.form.get("include_spec") == "on"

    # --- optional: DeepSeek smart fill for weight / unit price ----------
    ai_stats = None
    ai_log = []
    if request.form.get("use_ai") == "on":
        api_key = request.form.get("api_key", "").strip() or app_config.get_api_key()
        if request.form.get("save_key") == "on" and api_key:
            app_config.save_config({"api_key": api_key})
        if not api_key:
            flash("已勾选 DeepSeek 智能填写，但未提供 API Key（或未设置环境变量 DEEPSEEK_API_KEY）。", "error")
            return redirect(url_for("index"))
        try:
            ai_stats = enrich_parts(parts, api_key, on_status=ai_log.append)
            warnings.extend(ai_log)
        except Exception as e:
            flash(f"DeepSeek 调用失败：{e}", "error")
            return redirect(url_for("index"))

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
        ai_stats=ai_stats,
        trans_stats=trans_stats,
    )


@app.route("/download/<path:file_name>")
def download(file_name):
    target = (GENERATED_DIR / file_name).resolve()
    if not target.is_file() or target.parent != GENERATED_DIR:
        abort(404)
    return send_file(target, as_attachment=True, download_name=target.name)


if __name__ == "__main__":
    app.run(debug=True)