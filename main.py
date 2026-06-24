#!/usr/bin/env python
"""Unified CLI entry point for the ppt-skill framework (v2.0).

Builds HCG-Slide-Design-System v1.0 decks from a declarative client config.
Two render targets — PPTX (core/designer.py) and HTML (core/html_renderer.py).

Usage:
    python main.py --list
    python main.py --client lotte_chemical              # default: --both (.pptx + .html)
    python main.py --client lotte_chemical --pptx       # only .pptx
    python main.py --client lotte_chemical --html       # only .html
    python main.py --client lotte_chemical --dry-run    # print normalized spec, no render
    python main.py --client lotte_chemical --validate   # validate only, no render
    python main.py --client lotte_chemical --out C:\\path\\deck.pptx
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
PROJECTS_DIR = ROOT / "projects"
PLANNING_JSON = ROOT / "skill_ppt_planning.json"
DESIGN_JSON = ROOT / "skill_ppt_design.json"


def _load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def list_clients():
    if not PROJECTS_DIR.exists():
        return []
    return sorted(p.stem for p in PROJECTS_DIR.glob("*.json"))


def main(argv=None):
    ap = argparse.ArgumentParser(prog="ppt-skill")
    ap.add_argument("--client", help="projects/<client>.json to render")
    ap.add_argument("--list", action="store_true", help="list available clients")
    ap.add_argument("--out", help="override output path (extension set per format)")
    ap.add_argument("--dry-run", action="store_true",
                    help="print normalized spec JSON, no render")
    ap.add_argument("--validate", action="store_true",
                    help="validate config + planner only, no render")
    ap.add_argument("--pptx", action="store_true", help="render .pptx only")
    ap.add_argument("--html", action="store_true", help="render .html only")
    ap.add_argument("--both", action="store_true",
                    help="render both .pptx and .html (default)")
    args = ap.parse_args(argv)

    if args.list:
        for name in list_clients():
            print(name)
        return 0

    if not args.client:
        print("[error] --client is required (or use --list)", file=sys.stderr)
        return 2

    cfg_path = PROJECTS_DIR / f"{args.client}.json"
    if not cfg_path.exists():
        print(f"[error] config not found: {cfg_path}", file=sys.stderr)
        print(f"available clients: {', '.join(list_clients()) or '(none)'}",
              file=sys.stderr)
        return 2

    from core.planner import Planner
    from core import schema

    config = _load_json(cfg_path)
    planning_rules = _load_json(PLANNING_JSON) if PLANNING_JSON.exists() else {}

    planner = Planner(planning_rules)
    try:
        spec = planner.plan(config, out_override=args.out)
    except schema.ConfigError as exc:
        print(f"[config error] {exc}", file=sys.stderr)
        return 2

    for warning in planner.warnings:
        print(f"[warn] {warning}", file=sys.stderr)

    if args.validate:
        print(f"[ok] {args.client}: {len(spec['slides'])} slides valid")
        return 0

    if args.dry_run:
        print(json.dumps(spec, ensure_ascii=False, indent=2))
        return 0

    # which formats? default = both
    formats = []
    if args.pptx:
        formats.append("pptx")
    if args.html:
        formats.append("html")
    if args.both or not formats:
        formats = ["pptx", "html"]

    design_tokens = _load_json(DESIGN_JSON) if DESIGN_JSON.exists() else {}
    overrides = spec["meta"].get("colors")
    base = Path(spec["meta"]["out"])
    saved = []

    if "pptx" in formats:
        from core.designer import Designer
        path = str(base.with_suffix(".pptx"))
        Designer(out=path, design_tokens=design_tokens,
                 color_overrides=overrides).render(spec).save()
        saved.append(path)
    if "html" in formats:
        from core.html_renderer import HtmlRenderer
        path = str(base.with_suffix(".html"))
        HtmlRenderer(out=path, design_tokens=design_tokens,
                     color_overrides=overrides).render(spec)
        saved.append(path)

    print(f"[done] {args.client}: {', '.join(saved)}")
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")   # 콘솔 한글 깨짐 방지 (cp949 → utf-8)
    sys.stderr.reconfigure(encoding="utf-8")
    sys.exit(main())
