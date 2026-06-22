#!/usr/bin/env python
"""Unified CLI entry point for the ppt-skill framework.

Usage:
    python main.py --client lotte_chemical
    python main.py --list
    python main.py --client lotte_chemical --dry-run
    python main.py --client lotte_chemical --validate
    python main.py --client lotte_chemical --out C:\\path\\deck.pptx
"""
import argparse
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CONFIG_DIR = ROOT / "config"
PLANNING_JSON = ROOT / "skill_ppt_planning.json"
DESIGN_JSON = ROOT / "skill_ppt_design.json"


def _load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def list_clients():
    if not CONFIG_DIR.exists():
        return []
    return sorted(p.stem for p in CONFIG_DIR.glob("*.json"))


def main(argv=None):
    ap = argparse.ArgumentParser(prog="ppt-skill")
    ap.add_argument("--client", help="config/<client>.json to render")
    ap.add_argument("--list", action="store_true", help="list available clients")
    ap.add_argument("--out", help="override output .pptx path")
    ap.add_argument("--dry-run", action="store_true",
                    help="print normalized spec JSON, no render")
    ap.add_argument("--validate", action="store_true",
                    help="validate config + planner only, no render")
    args = ap.parse_args(argv)

    if args.list:
        for name in list_clients():
            print(name)
        return 0

    if not args.client:
        print("[error] --client is required (or use --list)", file=sys.stderr)
        return 2

    cfg_path = CONFIG_DIR / f"{args.client}.json"
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

    template = spec["meta"]["template"]
    if not Path(template).exists():
        print(f"[error] template not found: {template}", file=sys.stderr)
        return 2

    from core.designer import Designer
    design_tokens = _load_json(DESIGN_JSON) if DESIGN_JSON.exists() else {}
    designer = Designer(template, spec["meta"]["out"], theme=spec["meta"]["theme"],
                        design_tokens=design_tokens,
                        color_overrides=spec["meta"].get("colors"))
    designer.render(spec).save()
    print(f"[saved] {spec['meta']['out']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
