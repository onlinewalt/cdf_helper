"""CDF_helper - 报关清单生成器.

Usage:
    python main.py           启动 Web 界面（自动打开浏览器）
    python main.py generate --template <template.xlsx> --source <file1> [<file2> ...] \
        [--output-dir <folder>] [--vessel <name>] [--port <name>] [--date <date>] [--name-only]
"""

import argparse
import datetime
import sys
import threading
import time
import webbrowser
from pathlib import Path

from cdf_helper.generator import generate, sanitize_filename
from cdf_helper.parser import parse_sources, detect_vessel


def _ensure_utf8_stdio():
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

REPORT_LABEL = "报关清单"


def _parse_args(argv):
    p = argparse.ArgumentParser(
        prog="cdf_helper",
        description="根据备件来源 Excel 生成报关清单 (CDF)。",
    )
    sub = p.add_subparsers(dest="command")

    g = sub.add_parser("generate", help="命令行生成报关清单")
    g.add_argument("--template", default=None, help="报关清单模板 .xlsx 路径（不填则自动查找/询问）")
    g.add_argument("--source", nargs="+", default=None, help="一个或多个备件来源文件 (.xls/.xlsx)（不填则询问）")
    g.add_argument("--output-dir", default=None, help="保存目录（默认为当前目录下的 output/）")
    g.add_argument("--vessel", default=None, help="船名（默认从来源文件自动识别）")
    g.add_argument("--port", default=None, help="目的港")
    g.add_argument("--date", default=None, help="日期（默认今天），如 2026-08-20")
    g.add_argument("--name-only", action="store_true", help="备件名称列只填名称，不拼接规格")
    g.add_argument("--ai", action="store_true", help="用 DeepSeek 估算缺失的重量/单价")
    g.add_argument("--api-key", default=None, help="DeepSeek API Key（默认读环境变量 DEEPSEEK_API_KEY 或 config.json）")
    return p.parse_args(argv)


def _ask(question, default):
    if default:
        return input(f"{question}（默认: {default}）：").strip() or default
    return input(f"{question}: ").strip()


def _discover_files():
    """List spreadsheet files in cwd, excluding generated CDF outputs."""
    files = sorted(Path.cwd().glob("*.xls*"))
    return [f for f in files if not f.name.endswith("报关清单.xlsx") and "output" not in f.parts]


def _pick_template():
    candidates = [f for f in _discover_files() if f.suffix == ".xlsx"]
    default = None
    for f in candidates:
        if "报关清单" in f.name:
            default = f
            break
    if not default and candidates:
        default = candidates[0]
    while True:
        ans = _ask("请输入报关清单模板路径 (.xlsx)", str(default) if default else None)
        if ans and Path(ans).exists():
            return Path(ans)
        print(f"找不到文件: {ans}，请重新输入")


def _pick_sources(template=None):
    candidates = [f for f in _discover_files() if f != template]
    print("当前文件夹内可用的备件来源文件：")
    for i, f in enumerate(candidates, 1):
        print(f"  [{i}] {f.name}")
    while True:
        ans = input("请选择来源文件（可多个，用逗号或空格分隔，例如: 1 2 3；或直接输入路径）: ").strip()
        if not ans:
            continue
        if ans.startswith('['):
            continue
        chosen = []
        ok = True
        for token in ans.replace("，", " ").replace(",", " ").split():
            if token.isdigit():
                idx = int(token)
                if 1 <= idx <= len(candidates):
                    chosen.append(candidates[idx - 1])
                    continue
                ok = False
            else:
                p = Path(token.strip('"'))
                if p.exists():
                    chosen.append(p)
                    continue
                ok = False
        if not ok or not chosen:
            print("输入无效，请重试")
            continue
        return chosen


def _serve():
    """Start the Flask web app and open it in the default browser."""
    import socket

    import webapp

    def _free_port(start=5000, tries=20):
        for port in range(start, start + tries):
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                try:
                    s.bind(("127.0.0.1", port))
                    return port
                except OSError:
                    continue
        return start

    port = _free_port()
    url = f"http://127.0.0.1:{port}"
    print(f"正在启动 CDF Helper Web 界面：{url}")
    print("按 Ctrl+C 停止服务。")

    def open_browser():
        time.sleep(1.2)
        webbrowser.open(url)

    threading.Thread(target=open_browser, daemon=True).start()
    webapp.app.run(host="127.0.0.1", port=port, debug=False)


def main(argv=None):
    _ensure_utf8_stdio()
    args = _parse_args(argv if argv is not None else sys.argv[1:])

    if args.command is None:
        _serve()
        return 0
    if args.command != "generate":
        print(f"未知命令: {args.command}")
        return 1

    template = Path(args.template) if args.template else _pick_template()
    if not template.exists():
        print(f"错误：找不到模板文件 {template}")
        return 1

    if args.source:
        sources = [Path(s) for s in args.source]
    else:
        sources = _pick_sources(template)
    for s in sources:
        if not s.exists():
            print(f"错误：找不到来源文件 {s}")
            return 1

    warnings = []

    def warn(msg):
        warnings.append(msg)

    print("正在解析备件来源文件 ...")
    try:
        parts = parse_sources(sources, warn=warn)
    except ValueError as e:
        print(f"错误：{e}")
        return 1

    for msg in warnings:
        print(f"  提示: {msg}")
    print(f"共解析到 {len(parts)} 条备件。")

    vessel = args.vessel or detect_vessel(sources)
    if not vessel:
        vessel = _ask("请输入船名", None)

    port = args.port or _ask("请输入目的港", "")
    date = args.date or _ask("请输入日期", datetime.date.today().isoformat())

    if args.ai:
        from cdf_helper import config as app_config
        from cdf_helper.ai import enrich_parts

        api_key = args.api_key or app_config.get_api_key()
        if not api_key:
            print("错误：使用 --ai 需要提供 API Key（--api-key 或环境变量 DEEPSEEK_API_KEY）。")
            return 1
        print("正在用 DeepSeek 估算缺失的重量/单价 ...")

        def ai_status(msg):
            print(f"  {msg}")

        ai_stats = enrich_parts(parts, api_key, on_status=ai_status)
        print(f"DeepSeek 完成：新增 {ai_stats['filled']} 条，失败 {ai_stats['errors']} 条。")

    output_dir = Path(args.output_dir) if args.output_dir else Path.cwd() / "output"
    output_name = f"{sanitize_filename(vessel)}-{sanitize_filename(port)}-{REPORT_LABEL}-{sanitize_filename(date)}.xlsx"

    print(f"正在生成报关清单 {output_name} ...")
    out = generate(
        template_path=template,
        parts=parts,
        vessel_name=vessel,
        output_dir=output_dir,
        output_name=output_name,
        include_spec=not args.name_only,
    )
    print(f"完成：{out}  （共 {len(parts)} 条备件）")
    return 0


if __name__ == "__main__":
    sys.exit(main())