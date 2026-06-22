# -*- coding: utf-8 -*-
import json, collections
d = json.load(open(r"C:\Users\cgpar\ppt-skill\final_analysis.json", encoding="utf-8"))
out = []


def dump(s, ind):
    txt = " | ".join(p["text"] for p in s.get("text", []) if p["text"]) if s.get("text") else ""
    line = f"{'  '*ind}- {s['name']} [{s['shape_type']}] x={s['x_cm']} y={s['y_cm']} w={s['w_cm']} h={s['h_cm']} fill={s.get('fill')} line={s.get('line')}"
    out.append(line)
    if txt:
        out.append(f"{'  '*ind}    TXT: {txt[:90]}")
        for p in s.get("text", []):
            for r in p.get("runs", []):
                if r["text"].strip():
                    f = r["font"]
                    out.append(f"{'  '*ind}      run '{r['text'][:25]}' f={f.get('name')} sz={f.get('size')} b={f.get('bold')} c={f.get('color')} align={p.get('align')}")
    for c in s.get("children", []):
        dump(c, ind + 1)


# Full TOC master = slide 1
out.append("## SLIDE 1 = MASTER TOC (full group recursion)")
for s in d["slides"][1]["shapes"]:
    dump(s, 0)

out.append("\n## SLIDE 17 TOC (2nd section highlighted)")
for s in d["slides"][17]["shapes"]:
    dump(s, 0)

# color palette frequency across all fills + fonts
out.append("\n## COLOR PALETTE (fill + font color frequency)")
fillc = collections.Counter()
fontc = collections.Counter()


def scan(s):
    if s.get("fill") and s["fill"].get("color"):
        fillc[s["fill"]["color"]] += 1
    for p in s.get("text", []):
        for r in p.get("runs", []):
            c = r["font"].get("color")
            if c:
                fontc[c] += 1
    for c in s.get("children", []):
        scan(c)


for sl in d["slides"]:
    for s in sl["shapes"]:
        scan(s)
out.append("FILL colors (top 15):")
for c, n in fillc.most_common(15):
    out.append(f"  #{c}: {n}")
out.append("FONT colors (top 15):")
for c, n in fontc.most_common(15):
    out.append(f"  #{c}: {n}")

# font sizes
out.append("\n## FONT NAME/SIZE frequency")
fn = collections.Counter(); fs = collections.Counter()


def scanf(s):
    for p in s.get("text", []):
        for r in p.get("runs", []):
            if r["text"].strip():
                fn[r["font"].get("name")] += 1
                fs[r["font"].get("size")] += 1
    for c in s.get("children", []):
        scanf(c)


for sl in d["slides"]:
    for s in sl["shapes"]:
        scanf(s)
out.append(f"fonts: {dict(fn.most_common(8))}")
out.append(f"sizes: {dict(fs.most_common(12))}")

# a clean 3-card body slide: slide 7 detail
out.append("\n## SLIDE 7 'Why HCG' FULL DUMP (3-card example)")
for s in d["slides"][7]["shapes"]:
    dump(s, 0)

open(r"C:\Users\cgpar\ppt-skill\STRUCTURE_REPORT2.md", "w", encoding="utf-8").write("\n".join(out))
print("ok", len(out))
