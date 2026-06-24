# -*- coding: utf-8 -*-
"""HCG slide engine v2.0 — HTML renderer.

Emits a self-contained HTML deck implementing HCG-Slide-Design-System v1.0:
16:9 1280x720 px slides, Pretendard, single main BLUE palette, the rigid
7-element skeleton, conclusion titles, charts (Chart.js), source-note always.

Pure-Python (no python-pptx) — shares tokens with core/designer.py via
core/tokens.py. Pretendard + Chart.js load from CDN, so open the .html with a
network connection (or self-host the assets for offline use).
"""

import html
import json

from core.tokens import (PALETTE, TYPO, CONTENT, GUTTER, BLOCK_GAP, ROUNDED,
                          SERIES_PALETTE, ROMAN, apply_design_tokens, segments)

PRETENDARD_CDN = ("https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/"
                  "dist/web/static/pretendard.css")
CHARTJS_CDN = "https://cdn.jsdelivr.net/npm/chart.js@4"


def _esc(s):
    return html.escape(str(s), quote=True)


def _rich(text, base="ink", emph="primary"):
    """**emphasis** -> colored spans."""
    out = []
    for chunk, is_emph in segments(text):
        out.append(f'<span style="color:var(--{emph if is_emph else base})">'
                   f'{_esc(chunk)}</span>')
    return "".join(out)


def _role(role):
    px, wt, lh = TYPO[role]
    return f"font-size:{px}px;font-weight:{wt};line-height:{lh}"


def _abs(x, y, w, h):
    return f"position:absolute;left:{x}px;top:{y}px;width:{w}px;height:{h}px"


class HtmlRenderer:
    def __init__(self, out=None, design_tokens=None, color_overrides=None,
                 **_ignored):
        if design_tokens or color_overrides:
            apply_design_tokens(design_tokens or {}, color_overrides)
        self.out = out
        self.page = 0
        self.count = 0
        self._charts = []   # (canvas_id, chart_spec)

    # ── skeleton ─────────────────────────────────────────────
    def _skeleton(self, title=None, kicker=None, chapter=None, source=None,
                  page=None):
        h = []
        if kicker:
            base = "ink" if "**" in kicker else "primary"
            h.append(f'<div class="kicker" style="{_abs(56,28,760,28)};'
                     f'{_role("kicker")}">{_rich(kicker, base=base)}</div>')
        if chapter:
            h.append(f'<div class="chapter" style="{_abs(1074,28,150,26)};'
                     f'{_role("chapter-indicator")}">{_esc(chapter)}</div>')
        if title is not None:
            h.append(f'<div class="title" style="{_abs(56,64,1168,64)};'
                     f'{_role("slide-title")}">{_rich(title)}</div>')
            h.append(f'<div class="rule" style="{_abs(56,136,1168,1)}"></div>')
        if source:
            txt = source if str(source).strip().startswith("※") else "※ Source : " + source
            h.append(f'<div class="source" style="{_abs(56,678,980,16)};'
                     f'{_role("source-note")}">{_esc(txt)}</div>')
        if page:
            h.append(f'<div class="pageno" style="{_abs(1024,676,200,18)};'
                     f'{_role("page-number")}">{page} &nbsp; '
                     f'<span class="mark">HCG</span></div>')
        return "".join(h)

    # ── content builders -> inner HTML of .content-area ──────
    def _bullets(self, items, role="body"):
        rows = []
        for it in items:
            text = it if isinstance(it, str) else it.get("text", "")
            rows.append(f'<li class="bi" style="{_role(role)}">'
                        f'<span class="m">■</span>{_esc(text)}</li>')
            if isinstance(it, dict):
                for sub in it.get("sub", []) or []:
                    rows.append(f'<li class="bi sub" style="{_role("body-sm")}">'
                                f'<span class="m sub">–</span>{_esc(sub)}</li>')
        return f'<ul class="bullets">{"".join(rows)}</ul>'

    def build_bullets(self, s):
        return self._bullets(s["items"])

    def build_cards(self, s):
        items = s["items"]
        n = len(items)
        per_row = n if n <= 4 else (n + 1) // 2
        cards = []
        for it in items:
            header = it.get("header", "") if isinstance(it, dict) else ""
            head = (f'<div class="block-header">{_esc(header)}</div>'
                    if header else "")
            if isinstance(it, dict) and it.get("bullets"):
                body = self._bullets(it["bullets"], role="body-sm")
            else:
                body = (f'<div style="{_role("body-sm")};color:var(--ink-slate)">'
                        f'{_esc(it.get("body","") if isinstance(it,dict) else it)}</div>')
            cards.append(f'<div class="card">{head}<div class="card-body">{body}</div></div>')
        return (f'<div class="grid" style="grid-template-columns:repeat({per_row},1fr);'
                f'gap:{BLOCK_GAP}px {GUTTER}px">{"".join(cards)}</div>')

    def build_columns(self, s):
        cols = s["columns"]
        out = []
        for col in cols:
            dark = "dark" if col.get("dark") else ""
            hdr = f'<div class="col-header {dark}">{_esc(col.get("header",""))}</div>'
            if col.get("bullets"):
                body = self._bullets(col["bullets"])
            else:
                body = (f'<div class="col-body">{_esc(col.get("body",""))}</div>')
            out.append(f'<div class="col">{hdr}{body}</div>')
        return (f'<div class="grid" style="grid-template-columns:repeat({len(cols)},1fr);'
                f'gap:{GUTTER}px">{"".join(out)}</div>')

    def build_compare(self, s):
        def column(rows, header, dark):
            cls = "col-header dark" if dark else "col-header"
            cells = "".join(
                f'<div class="cmp-cell {"tb" if dark else ""}" '
                f'style="{_role("body-sm")}">{_esc(r)}</div>' for r in rows)
            return (f'<div class="cmp-col"><div class="{cls}">{_esc(header)}</div>'
                    f'<div class="cmp-list">{cells}</div></div>')
        left = column(s["left"], s.get("left_header", "As-Is"), False)
        right = column(s["right"], s.get("right_header", "To-Be"), True)
        take = (f'<div class="chip">{_esc(s["takeaway"])}</div>'
                if s.get("takeaway") else "")
        return (f'<div class="cmp">{left}{right}</div>{take}')

    def build_kpi(self, s):
        cells = []
        for it in s["items"]:
            unit = (f'<span class="kpi-unit">{_esc(it.get("unit",""))}</span>'
                    if it.get("unit") else "")
            cells.append(
                f'<div class="kpi"><div class="kpi-num">{_esc(it.get("number",""))}'
                f'{unit}</div><div class="kpi-label">{_esc(it.get("label",""))}</div></div>')
        return f'<div class="kpi-row">{"".join(cells)}</div>'

    def build_process(self, s):
        steps = s["steps"]
        parts = []
        for i, st in enumerate(steps):
            desc = (f'<div class="proc-desc">{_esc(st.get("desc",""))}</div>'
                    if st.get("desc") else "")
            parts.append(f'<div class="proc-step"><div class="proc-node">'
                         f'{_esc(st.get("label",""))}</div>{desc}</div>')
            if i < len(steps) - 1:
                parts.append('<div class="proc-arrow">▶</div>')
        return f'<div class="proc">{"".join(parts)}</div>'

    def build_matrix(self, s):
        cols, rows = s["columns"], s["rows"]
        head = "".join(f'<th>{_esc(c)}</th>' for c in cols)
        body = []
        for r in rows:
            cells = r.get("cells", []) if isinstance(r, dict) else list(r)
            group = r.get("group", "") if isinstance(r, dict) else ""
            tds = ""
            if group:
                tds += f'<td class="grp">{_esc(group)}</td>'
            tds += "".join(f"<td>{_esc(c)}</td>" for c in cells)
            body.append(f"<tr>{tds}</tr>")
        take = (f'<div class="chip">{_esc(s["takeaway"])}</div>'
                if s.get("takeaway") else "")
        return (f'<table class="matrix"><thead><tr>{head}</tr></thead>'
                f'<tbody>{"".join(body)}</tbody></table>{take}')

    def build_table(self, s):
        headers, rows = s["headers"], s["rows"]
        emph = s.get("emphasis_row")
        head = "".join(f"<th>{_esc(c)}</th>" for c in headers)
        body = []
        for i, r in enumerate(rows):
            cls = ' class="emph"' if (emph is not None and i == emph) else ""
            body.append(f"<tr{cls}>" + "".join(f"<td>{_esc(c)}</td>" for c in r) + "</tr>")
        return (f'<table class="data-table"><thead><tr>{head}</tr></thead>'
                f'<tbody>{"".join(body)}</tbody></table>')

    def build_chart(self, s):
        cid = f"chart{len(self._charts)}"
        self._charts.append((cid, s["chart"]))
        return f'<div class="chart-wrap"><canvas id="{cid}"></canvas></div>'

    # ── structural slides (full HTML) ────────────────────────
    def _cover(self, s):
        conf = ('<div class="cover-conf">Strictly Confidential</div>'
                if s.get("confidential", True) else "")
        sub = (f'<div class="cover-sub" style="{_abs(160,470,960,30)};'
               f'{_role("cover-subtitle")}">{_esc(s["subtitle"])}</div>'
               if s.get("subtitle") else "")
        date = (f'<div class="cover-date" style="{_abs(160,540,960,24)};'
                f'{_role("cover-date")}">{_esc(s["date"])}</div>'
                if s.get("date") else "")
        legal = s.get("legal") or ("본 제안서는 ㈜휴먼컨설팅그룹의 편집저작물로서 "
                                   "관련 법령에 의해 보호됩니다.")
        return (
            '<div class="cover-band"></div>' + conf +
            f'<div class="cover-title" style="{_abs(160,392,960,56)};'
            f'{_role("cover-title")}">{_rich(s["title"])}</div>' + sub + date +
            f'<div class="cover-logo" style="{_abs(560,612,160,48)};'
            f'{_role("section-header")}">HCG</div>'
            f'<div class="cover-legal" style="{_abs(160,676,960,28)};'
            f'{_role("source-note")}">{_esc(legal)}</div>')

    def _section(self, s):
        sub = (f'<div class="sec-sub" style="{_abs(300,372,884,60)};'
               f'{_role("cover-subtitle")}">{_esc(s["sub"])}</div>'
               if s.get("sub") else "")
        return (
            '<div class="sec-band"></div>'
            f'<div class="sec-num" style="{_abs(96,250,300,120)};'
            f'{_role("kpi-number")}">{_esc(s.get("num",""))}</div>'
            f'<div class="sec-title" style="{_abs(300,300,884,56)};'
            f'{_role("cover-title")}">{_rich(s["title"])}</div>' + sub)

    def _end(self, s):
        return (
            f'<div class="end-msg" style="{_role("cover-title")}">'
            f'{_esc(s.get("message","Thank You"))}</div>'
            f'<div class="end-org" style="{_role("cover-subtitle")}">㈜휴먼컨설팅그룹</div>')

    _CONTENT = {"bullets": "build_bullets", "cards": "build_cards",
                "columns": "build_columns", "compare": "build_compare",
                "kpi": "build_kpi", "process": "build_process",
                "matrix": "build_matrix", "chart": "build_chart",
                "table": "build_table"}

    # ── render ───────────────────────────────────────────────
    def render(self, spec):
        slides = []
        for s in spec.get("slides", []):
            t = s.get("type")
            self.count += 1
            if t == "cover":
                inner, cls = self._cover(s), "slide cover"
            elif t == "section":
                inner, cls = self._section(s), "slide section"
            elif t == "end":
                inner, cls = self._end(s), "slide end"
            elif t == "toc":
                self.page += 1
                inner = self._toc(s) + self._skeleton(source=s.get("source"),
                                                       page=self.page)
                cls = "slide toc"
            elif t in self._CONTENT:
                self.page += 1
                body = getattr(self, self._CONTENT[t])(s)
                frame = self._skeleton(title=s.get("title"), kicker=s.get("kicker"),
                                       chapter=s.get("chapter"),
                                       source=s.get("source"), page=self.page)
                inner = (frame + f'<div class="content-area" '
                         f'style="{_abs(**CONTENT)}">{body}</div>')
                cls = "slide content"
            else:
                self.count -= 1
                continue
            slides.append(f'<section class="{cls}">{inner}</section>')
        doc = self._document("".join(slides))
        if self.out:
            with open(self.out, "w", encoding="utf-8") as f:
                f.write(doc)
            print(f"[saved html] {self.out}")
        return doc

    def _toc(self, s):
        items, current = s["items"], s.get("current")
        rows = []
        n = len(items)
        row_h = min(86, 512 / n)
        top = 152 + (512 - row_h * n) / 2
        for i, name in enumerate(items):
            active = "active" if current == i else ""
            y = top + i * row_h
            rows.append(
                f'<div class="toc-row {active}" style="{_abs(320,int(y),640,int(row_h))}">'
                f'<span class="toc-num">{ROMAN[i] if i < len(ROMAN) else i+1}</span>'
                f'<span class="toc-name" style="{_role("section-header")}">{_esc(name)}</span>'
                f'</div>')
        title = (f'<div class="title" style="{_abs(56,64,1168,64)};'
                 f'{_role("slide-title")}">{_esc(s.get("title","목차"))}</div>'
                 f'<div class="rule" style="{_abs(56,136,1168,1)}"></div>')
        return title + "".join(rows)

    def _chart_scripts(self):
        if not self._charts:
            return ""
        pal = [PALETTE[c] for c in SERIES_PALETTE]
        blocks = []
        for cid, spec in self._charts:
            ctype = spec.get("type", "column")
            single = ctype in ("pie", "donut", "doughnut")
            js_type = {"column": "bar", "stacked-column": "bar",
                       "stacked-bar": "bar", "donut": "doughnut",
                       "area": "line"}.get(ctype, ctype)
            datasets = []
            for i, ser in enumerate(spec.get("series", [])):
                if single:
                    datasets.append({"label": ser.get("name", ""),
                                     "data": ser.get("data", []),
                                     "backgroundColor": pal})
                elif ctype in ("line", "area"):
                    datasets.append({"label": ser.get("name", ""),
                                     "data": ser.get("data", []),
                                     "borderColor": pal[i % len(pal)],
                                     "backgroundColor": pal[i % len(pal)],
                                     "fill": ctype == "area", "tension": 0.3})
                else:
                    datasets.append({"label": ser.get("name", ""),
                                     "data": ser.get("data", []),
                                     "backgroundColor": pal[i % len(pal)]})
            indexaxis = "'y'" if ctype in ("bar", "stacked-bar") else "'x'"
            stacked = ctype in ("stacked-bar", "stacked-column")
            cfg = {
                "type": js_type,
                "data": {"labels": spec.get("labels", []), "datasets": datasets},
                "options": {
                    "responsive": True, "maintainAspectRatio": False,
                    "indexAxis": "AXIS",
                    "plugins": {"legend": {"display": len(datasets) > 1 or single,
                                           "position": "bottom"}},
                    "scales": ({} if single else
                               {"x": {"stacked": stacked}, "y": {"stacked": stacked}}),
                },
            }
            cfg_json = json.dumps(cfg, ensure_ascii=False).replace('"AXIS"', indexaxis)
            blocks.append(f'new Chart(document.getElementById("{cid}"),{cfg_json});')
        return (f'<script src="{CHARTJS_CDN}"></script>\n<script>'
                f'Chart.defaults.font.family="Pretendard";'
                f'Chart.defaults.color="{PALETTE["text-muted"]}";'
                f'{"".join(blocks)}</script>')

    def _document(self, body):
        cssvars = ";".join(f"--{k}:{v}" for k, v in PALETTE.items())
        return f"""<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>HCG Deck</title>
<link rel="stylesheet" href="{PRETENDARD_CDN}">
<style>
:root{{{cssvars}}}
*{{margin:0;padding:0;box-sizing:border-box}}
body{{background:#3a3a3a;font-family:Pretendard,system-ui,sans-serif;
  -webkit-font-smoothing:antialiased;padding:24px;display:flex;
  flex-direction:column;align-items:center;gap:24px}}
.slide{{position:relative;width:1280px;height:720px;background:var(--surface);
  overflow:hidden;box-shadow:0 6px 28px rgba(0,0,0,.35);flex:0 0 auto}}
.kicker{{display:flex;align-items:center}}
.chapter{{display:flex;align-items:center;justify-content:center;
  color:var(--primary-steel);background:var(--surface);
  border:1px solid var(--divider);border-radius:{ROUNDED["md"]}px}}
.title{{display:flex;align-items:center;color:var(--ink)}}
.rule{{background:var(--divider)}}
.source{{display:flex;align-items:center;color:var(--text-muted)}}
.pageno{{display:flex;align-items:center;justify-content:flex-end;color:var(--text-muted)}}
.pageno .mark{{color:var(--hcg-maroon);font-weight:700;margin-left:6px}}
.content-area{{display:flex;flex-direction:column}}
.bullets{{list-style:none;display:flex;flex-direction:column;gap:8px;padding-top:8px}}
.bi{{display:flex;gap:8px;color:var(--ink)}}
.bi .m{{color:var(--primary);font-weight:700}}
.bi.sub{{margin-left:26px;color:var(--ink-slate)}}
.bi.sub .m{{color:var(--text-muted);font-weight:400}}
.grid{{display:grid;flex:1;min-height:0}}
.card{{background:var(--surface);border:1px solid var(--divider);
  border-radius:{ROUNDED["card"]}px;overflow:hidden;display:flex;flex-direction:column}}
.block-header{{background:var(--divider-soft);color:var(--ink);text-align:center;
  padding:6px 10px;border-radius:{ROUNDED["sm"]}px;{_role("block-header")}}}
.card-body{{padding:10px 12px;flex:1;{_role("body-sm")};color:var(--ink-slate)}}
.col{{display:flex;flex-direction:column;gap:12px}}
.col-header{{background:var(--primary-soft);color:var(--ink);text-align:center;
  padding:8px 12px;border-radius:{ROUNDED["sm"]}px;{_role("section-header")}}}
.col-header.dark{{background:var(--primary-deep);color:var(--surface)}}
.col-body{{background:var(--surface);border:1px solid var(--divider);
  border-radius:{ROUNDED["card"]}px;padding:12px 14px;flex:1;{_role("body")};color:var(--ink-slate)}}
.cmp{{display:grid;grid-template-columns:1fr 1fr;gap:{GUTTER}px;flex:1;min-height:0}}
.cmp-col{{display:flex;flex-direction:column;gap:12px;min-height:0}}
.cmp-list{{display:flex;flex-direction:column;gap:{BLOCK_GAP}px;flex:1}}
.cmp-cell{{background:var(--surface);border:1px solid var(--divider);
  border-radius:{ROUNDED["card"]}px;padding:10px 12px;flex:1;display:flex;align-items:center;color:var(--ink)}}
.cmp-cell.tb{{background:var(--tint-blue-4);border-color:var(--tint-blue-3)}}
.chip{{align-self:center;background:var(--accent-coral);color:#fff;
  border-radius:{ROUNDED["chip"]}px;padding:3px 16px;margin-top:12px;{_role("body-sm-strong")}}}
.kpi-row{{display:flex;gap:{GUTTER}px;flex:1;align-items:center}}
.kpi{{flex:1;background:var(--tint-blue-4);border:1px solid var(--tint-blue-3);
  border-radius:{ROUNDED["card"]}px;padding:24px;text-align:center}}
.kpi-num{{color:var(--primary);{_role("kpi-number")}}}
.kpi-unit{{color:var(--text-muted);{_role("kpi-unit")};margin-left:4px}}
.kpi-label{{color:var(--ink-slate);{_role("caption")};margin-top:10px}}
.proc{{display:flex;align-items:flex-start;gap:8px;flex:1;padding-top:40px}}
.proc-step{{flex:1;display:flex;flex-direction:column;gap:10px;align-items:center}}
.proc-node{{width:100%;background:var(--tint-blue-3);color:var(--ink);
  border-radius:{ROUNDED["card"]}px;padding:18px 8px;text-align:center;{_role("body-sm-strong")}}}
.proc-desc{{{_role("body-sm")};color:var(--ink-slate);text-align:center}}
.proc-arrow{{color:var(--text-muted);align-self:center;padding-top:18px;font-size:22px}}
table{{border-collapse:collapse;width:100%}}
.matrix,.data-table{{flex:0 0 auto;{_role("body-sm")}}}
.matrix th,.data-table th{{background:var(--tint-blue-1);color:var(--ink);
  {_role("body-sm-strong")};padding:10px;border:1px solid var(--surface)}}
.matrix td,.data-table td{{padding:8px 10px;border:1px solid var(--divider);color:var(--ink)}}
.data-table tbody tr:nth-child(even){{background:var(--surface-gray)}}
.data-table tr.emph{{background:var(--tint-blue-2)}}
.data-table tr.emph td{{font-weight:700}}
.matrix td.grp{{background:var(--tint-blue-4);color:var(--primary);
  {_role("body-sm-strong")};text-align:center}}
.chart-wrap{{flex:1;position:relative;min-height:0}}
.cover-band{{position:absolute;left:0;top:0;width:1280px;height:300px;
  background:var(--primary);border-bottom:6px solid var(--primary-deep)}}
.cover-conf{{position:absolute;left:56px;top:20px;color:#fff;font-style:italic;{_role("caption")}}}
.cover-title{{display:flex;align-items:center;justify-content:center;text-align:center;color:var(--ink)}}
.cover-sub{{text-align:center;color:var(--primary-steel)}}
.cover-date{{text-align:center;color:var(--text-muted)}}
.cover-logo{{display:flex;align-items:center;justify-content:center;color:var(--hcg-maroon)}}
.cover-legal{{text-align:center;color:var(--text-muted)}}
.sec-band{{position:absolute;left:0;top:240px;width:1280px;height:240px;background:var(--tint-blue-4)}}
.sec-num{{display:flex;align-items:center;color:var(--primary)}}
.sec-title{{display:flex;align-items:center;color:var(--ink)}}
.sec-sub{{display:flex;align-items:center;color:var(--ink-slate)}}
.end{{background:var(--primary);display:flex;flex-direction:column;
  align-items:center;justify-content:center;gap:16px}}
.end-msg{{color:#fff}}
.end-org{{color:var(--tint-blue-2)}}
.toc-row{{position:absolute;display:flex;align-items:center;gap:24px}}
.toc-num{{display:flex;align-items:center;justify-content:center;width:64px;height:52px;
  background:var(--tint-blue-2);color:#fff;border-radius:{ROUNDED["sm"]}px;{_role("section-header")}}}
.toc-row.active .toc-num{{background:var(--primary)}}
.toc-name{{color:var(--ink)}}
.toc-row.active .toc-name{{color:var(--primary);font-weight:700}}
.toc-row.active{{outline:1.5px solid var(--primary);outline-offset:6px;border-radius:{ROUNDED["sm"]}px}}
</style></head>
<body>
{body}
{self._chart_scripts()}
</body></html>"""

    def save(self, out=None):
        # symmetry with DeckEngine.save(); render() already wrote if self.out set
        path = out or self.out
        return path
