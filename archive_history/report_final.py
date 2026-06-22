# -*- coding: utf-8 -*-
"""Phase 1 report: clean UTF-8 summary of key structural patterns."""
import json

d = json.load(open(r"C:\Users\cgpar\ppt-skill\final_analysis.json", encoding="utf-8"))
out = []
W = d["slide_w_cm"]; H = d["slide_h_cm"]
out.append(f"# FINAL PPTX STRUCTURE  ({W} x {H} cm, {d['n_slides']} slides)\n")


def title_of(sl):
    for s in sl["shapes"]:
        if s.get("text"):
            t = s["text"][0]["text"].strip()
            if t:
                return t
    return "(none)"


# 1. layouts
from collections import Counter
out.append("## LAYOUTS")
for k, v in Counter(s["layout"] for s in d["slides"]).items():
    out.append(f"- '{k}': {v}")
out.append("")

# 2. find TOC slide (title or many numbered short items)
out.append("## TOC / 목차 DETECTION")
for sl in d["slides"]:
    t = title_of(sl)
    if "목차" in t or "Contents" in t or "CONTENTS" in t or "Agenda" in t.lower():
        out.append(f"### slide {sl['idx']} title='{t}' shapes={sl['n_shapes']}")
for sl in d["slides"]:
    if "목차" in title_of(sl):
        out.append(f"\n>>> FULL TOC SHAPE DUMP slide {sl['idx']} <<<")
        for s in sl["shapes"]:
            txt = " | ".join(p["text"] for p in s.get("text", []) if p["text"]) if s.get("text") else ""
            out.append(f"  - {s['name']} [{s['shape_type']}] x={s['x_cm']} y={s['y_cm']} w={s['w_cm']} h={s['h_cm']} fill={s.get('fill')} line={s.get('line')}")
            if txt:
                out.append(f"      text: {txt[:80]}")
                for p in s.get("text", []):
                    for r in p["runs"]:
                        f = r["font"]
                        out.append(f"        run '{r['text'][:30]}' font={f.get('name')} sz={f.get('size')} b={f.get('bold')} c={f.get('color')}")
out.append("")

# 3. divider slides (<=2 shapes, non-empty)
out.append("## SECTION DIVIDER SLIDES (1-3 shapes)")
for sl in d["slides"]:
    if 1 <= sl["n_shapes"] <= 3:
        out.append(f"### slide {sl['idx']} title='{title_of(sl)}' shapes={sl['n_shapes']} layout='{sl['layout']}'")
        for s in sl["shapes"]:
            txt = s["text"][0]["text"] if s.get("text") else ""
            out.append(f"  - {s['name']} x={s['x_cm']} y={s['y_cm']} w={s['w_cm']} h={s['h_cm']} fill={s.get('fill')} txt='{txt[:40]}'")
out.append("")

# 4. richest grid slides
out.append("## GRID-ARRANGEMENT SLIDES (>=3 shapes in a row)")
grid_slides = [sl for sl in d["slides"] if sl["grids"]]
# sort by best grid count
grid_slides.sort(key=lambda s: max((g["count"] for g in s["grids"]), default=0), reverse=True)
for sl in grid_slides[:8]:
    out.append(f"### slide {sl['idx']} title='{title_of(sl)}'")
    for g in sl["grids"]:
        out.append(f"  row y={g['y_cm']} count={g['count']} x={g['x_positions']} w={g['widths']}")
out.append("")

# 5. container detection: shapes whose bbox contains other shapes (box-in-box)
out.append("## CONTAINER (BOX-IN-BOX) DETECTION")


def contains(a, b):
    if None in (a["x_cm"], a["y_cm"], a["w_cm"], a["h_cm"], b["x_cm"], b["y_cm"], b["w_cm"], b["h_cm"]):
        return False
    return (a["x_cm"] <= b["x_cm"] + 0.2 and a["y_cm"] <= b["y_cm"] + 0.2 and
            a["x_cm"] + a["w_cm"] + 0.2 >= b["x_cm"] + b["w_cm"] and
            a["y_cm"] + a["h_cm"] + 0.2 >= b["y_cm"] + b["h_cm"] and
            (a["w_cm"] * a["h_cm"]) > (b["w_cm"] * b["h_cm"]) * 1.15)


for sl in d["slides"]:
    shapes = [s for s in sl["shapes"] if s["w_cm"]]
    pairs = []
    for a in shapes:
        if a.get("fill") and a["fill"].get("color"):  # only filled outer boxes
            inner = [b for b in shapes if b is not a and contains(a, b)]
            if inner:
                pairs.append((a, inner))
    if pairs:
        out.append(f"### slide {sl['idx']} title='{title_of(sl)}'  containers={len(pairs)}")
        for a, inner in pairs[:4]:
            out.append(f"  OUTER {a['name']} x={a['x_cm']} y={a['y_cm']} w={a['w_cm']} h={a['h_cm']} fill={a['fill']}")
            for b in inner[:5]:
                bt = b["text"][0]["text"][:30] if b.get("text") else ""
                out.append(f"     INNER {b['name']} x={b['x_cm']} y={b['y_cm']} w={b['w_cm']} h={b['h_cm']} txt='{bt}'")

open(r"C:\Users\cgpar\ppt-skill\STRUCTURE_REPORT.md", "w", encoding="utf-8").write("\n".join(out))
print("wrote STRUCTURE_REPORT.md", len(out), "lines")
