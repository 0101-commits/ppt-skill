# ppt-skill Unified Framework — Refactor Design

- **Date:** 2026-06-19
- **Status:** Approved (brainstorming) — pending implementation plan
- **Repo:** `C:\Users\cgpar\ppt-skill`
- **Author:** cgpark@e-hcg.com + Claude

## 1. Problem

The repo holds multiple per-client PPT generators with high redundancy:

- `auto_ppt.py` (1928 lines) — the de-facto master engine: XML primitives, shape
  builders, ~20 `build_*` slide builders, advanced diagram builders, `DeckEngine`
  with a 19-type `render(spec)` dispatch, a measured-coordinate block, and a
  `THEME`/`apply_theme` system. It also still carries a legacy hardcoded
  `build()` for the 롯데알미늄 deck plus a `__main__` CLI.
- `auto_ppt_kia.py` (486 lines), `lotte_chemical_ppt.py` (432 lines) — **imperative**
  clients that `import` builders from `auto_ppt` and construct slides directly.
  No spec JSON.
- `paradise_compare.py` (237 lines) — an analysis/QA tool (not a generator).
- `analyze_final.py`, `report2.py`, `report_final.py` — one-off analysis scripts.

The planning↔design separation already exists **as data** — `skill_ppt_planning.json`
(storyline, content rules) and `skill_ppt_design.json` (coords, colors, layout specs) —
but the **code does not mirror it**. Adding a client means writing a new imperative script.

## 2. Goal

A single, data-driven framework that cleanly separates **Planning** (content →
layout structuring) from **Design** (spec → pptx rendering). New clients are added
by writing a `config/<client>.json` file, not Python. The proven rendering engine is
**preserved**, not rewritten.

## 3. Decisions (locked during brainstorming)

| # | Decision | Choice | Rationale |
|---|----------|--------|-----------|
| D1 | Refactor depth | **Preserve engine, reorganize** | `auto_ppt.py` is tested; measured coords + XML helpers cost many iterations. Re-deriving risks regressions + reintroducing fixed bugs. |
| D2 | Planner intelligence | **Declarative** | config states each slide's `type` explicitly. Planner validates + structures; no inference. Predictable, debuggable. Matches today's spec-driven flow. |
| D3 | Migration scope | **Framework only** | Build planner/designer/main/schema + one representative sample config (`lotte_chemical.json`). Do NOT reverse-engineer kia/paradise/lotte_aluminium decks this pass. |
| D4 | Design-token consumption | **Partial** | designer consumes `skill_ppt_design.json` for colors/fonts/tokens + validation reference, but keeps proven Python coordinate constants. Full coord-from-JSON rewire deferred (would contradict D1). |

## 4. Target directory layout

```
ppt-skill/
├── main.py                 # CLI: --client <name> [--list] [--out] [--dry-run] [--validate]
├── core/
│   ├── __init__.py
│   ├── planner.py          # NEW. config content -> normalized spec. NO pptx import.
│   ├── designer.py         # PRESERVED engine. = old auto_ppt.py minus legacy build()+CLI.
│   └── schema.py           # NEW. slide-type registry, validation, error types.
├── config/
│   └── lotte_chemical.json # sample client = identity + content
├── skill_ppt_planning.json # EXISTING. planner reads.
├── skill_ppt_design.json   # EXISTING. designer reads (colors/fonts/tokens).
├── archive_history/
│   ├── README.md           # "frozen against pre-refactor API"
│   ├── auto_ppt_legacy.py  # snapshot of old monolith (safety copy)
│   ├── auto_ppt_kia.py
│   ├── lotte_chemical_ppt.py
│   ├── paradise_compare.py
│   ├── analyze_final.py
│   ├── report2.py
│   └── report_final.py
└── README.md               # updated usage
```

## 5. Module responsibilities

### 5.1 core/schema.py (new, ~120 lines)

- `SLIDE_TYPES`: registry mapping each of the **19 slide types** to its
  `{required: [...], optional: [...]}` key sets. Single source of truth, derived
  from the current `render()` dispatch.
  - Types: `cover`, `toc`, `section_agenda`, `blocks`, `container`, `overview`,
    `section`, `body_2col`, `body_single`, `body_process`, `overview_3col`,
    `diff_matrix`, `pain_point_categorized`, `approach_vs`, `process_roadmap`,
    `compare_table`, `appendix`, `demo_advanced`, `end`.
- `validate_config(cfg)` — checks `identity` + `content` present and well-formed.
- `validate_spec(spec)` — post-planner sanity check (types known, required keys present).
- Error classes: `ConfigError`, `UnknownSlideType`, `MissingField`, `TextOverflowWarning`.
- **Fixes a current bug:** `render()` today silently skips unknown types. Schema makes
  it a loud error naming the offending type and listing valid ones.

### 5.2 core/planner.py (new, ~200 lines — pure dict→dict, zero pptx)

- `Planner(planning_rules: dict)` — constructed with parsed `skill_ppt_planning.json`.
- `plan(config: dict) -> spec: dict`:
  1. `validate_config(config)`.
  2. Resolve storyline: order slides per `storyline_architecture`; assign section
     numbers (Ⅰ~Ⅳ) to section dividers.
  3. Inject/normalize: head-messages, default subtitle, `meta` block built from
     `identity` (theme, output path, template path).
  4. Declarative enforcement: every slide must carry an explicit `type`. Missing or
     unknown → `ConfigError`. No inference (D2).
  5. Emit normalized `spec = {meta: {...}, slides: [...]}` — exactly the shape
     `designer.render()` consumes.
- No `pptx` import anywhere. Independently unit-testable without PowerPoint.

### 5.3 core/designer.py (preserved engine — old auto_ppt.py)

- Copied **verbatim**: 12 XML primitives (`set_text_outline`, `add_shadow`,
  `set_transparency`, `set_gradient`, `set_char_spacing`, `set_text_anchor`,
  `_add_run`, `_set_font_xml`, `_set_line_spacing`, `_set_no_fill`, `_set_line_dark`,
  `set_autofit_shrink`/`fit_font_size`); ~15 shape builders (`add_item`, `add_textbox`,
  badges, connectors, `add_table`); 20 `build_*` slide builders; 3 advanced diagram
  builders (`create_toc_slide`, `add_structured_content_blocks`, `add_container_box`);
  `THEME` dict + `apply_theme`; measured coordinate block (`TITLE_Y`, `COL_L_X`,
  `ITEM_W`, `ITEM_Y0`, `ITEM_DY`, etc.); the `DeckEngine` class with its 19-type
  `render(spec)` dispatch and `save()`.
- **Removed / extracted out:**
  - Legacy `build()` (롯데알미늄 hardcode) → moved to `archive_history` as a frozen client.
  - `__main__` CLI block → moved to `main.py`.
  - Hardcoded template paths (`REAL`, `KIA_FINAL`) → no longer module constants; the
    template path is now passed in from config `identity.template`.
- Public surface: `Designer` is `DeckEngine` (alias or thin facade). Usage:
  `Designer(template, out, theme).render(spec).save()`.
- Design-JSON consumption (D4): loads `skill_ppt_design.json` to source colors/fonts/
  tokens and to allow `identity.colors`/`identity.fonts` overrides; coordinate constants
  remain in Python.
- Brand palettes (`KIA_*`, `PARADISE_*`, `HCG_*`) remain as named theme presets so
  `theme: "hcg"` / `theme: "paradise"` keep working; `identity.colors` layers on top.

### 5.4 main.py (new, ~120 lines)

- CLI entry point. Flags:
  - `--client <name>` (required) — loads `config/<name>.json`.
  - `--list` — list available client configs and exit.
  - `--out <path>` — override output path.
  - `--dry-run` — run planner only, print normalized spec JSON, no pptx.
  - `--validate` — validate config + planner, no render.
- Flow: load `config/<client>.json` + `skill_ppt_planning.json` + `skill_ppt_design.json`
  → `Planner.plan(config)` → `Designer(template, out, theme).render(spec).save()`.
- Friendly errors: missing config → message + available-client list.

## 6. Data flow

```
config/<client>.json
   ├─ identity{theme, template, colors, fonts, logo}
   └─ content{title, subtitle, date, slides[]}
                 │
      planner.plan()  ← reads skill_ppt_planning.json
                 │  (validate, order, section-number, head-message)
                 ▼
      normalized spec {meta, slides[]}
                 │
      designer.render(spec)  ← reads skill_ppt_design.json + identity
                 ▼
             output .pptx
```

## 7. config/&lt;client&gt;.json schema

```json
{
  "client": "lotte_chemical",
  "identity": {
    "display_name": "롯데케미칼",
    "theme": "hcg",
    "template": "templates/hcg_base.pptx",
    "colors":  {"primary": "#921F0B", "secondary": "#919191", "background": "#FFFFFF"},
    "fonts":   {"kr": "맑은 고딕", "en": "Arial"},
    "logo":    "assets/lotte_chemical.png"
  },
  "content": {
    "title": "...",
    "subtitle": "...",
    "date": "2025.05",
    "slides": [
      {"type": "cover", "...": "..."},
      {"type": "toc", "items": ["...", "..."]},
      {"type": "approach_vs",
       "left_title": "...", "left": ["..."],
       "right_title": "...", "right": ["..."]}
    ]
  }
}
```

- `identity.colors`, `identity.logo` are optional (theme preset supplies defaults).
- Planner maps `identity` → `spec.meta` and `content.slides` → `spec.slides`.
- The shipped sample covers ~6 slide types to demonstrate the full mapping
  (cover, toc, overview, approach_vs, diff_matrix, end).

## 8. Error handling (requirement #3)

| Condition | Behavior |
|-----------|----------|
| Missing config file | Wrapped `FileNotFoundError` + list of available clients |
| Missing template file | Clear error naming the missing path |
| Unknown slide `type` | `UnknownSlideType` naming bad type + listing the valid 19 (no silent skip) |
| Missing required field | `MissingField` naming slide index + field |
| Text overflow | Existing `fit_font_size()` / `set_autofit_shrink()` preserved; planner warns on over-limit text |

## 9. Testing

- **planner** — pure dict→dict unit tests: storyline ordering, section numbering,
  validation errors (missing type, unknown type, missing field).
- **designer** — smoke test: render the sample config, assert the `.pptx` opens and
  slide count matches expectation; no exception raised.
- **schema** — unit tests for bad-type and missing-field detection.

## 10. Archival

- `git mv` legacy scripts → `archive_history/` (git history preserved).
- Snapshot current `auto_ppt.py` → `archive_history/auto_ppt_legacy.py` before
  transforming it into `core/designer.py` (safety copy).
- `archive_history/README.md` notes the scripts are frozen against the pre-refactor API
  and will be migrated to config in a follow-up.

## 11. Out of scope (deferred)

- Reverse-engineering kia / paradise / lotte_aluminium real decks into config (D3).
- Full coordinate-from-JSON rewire of designer (D4).
- Inferential / heuristic layout auto-detection (D2).

## 12. Acceptance criteria

1. `python main.py --client lotte_chemical` produces a valid `.pptx` with no exception.
2. `python main.py --list` lists `lotte_chemical`.
3. `python main.py --client lotte_chemical --dry-run` prints normalized spec JSON
   without writing a pptx.
4. An unknown slide type in a config produces `UnknownSlideType`, not a silent skip.
5. `core/planner.py` imports no pptx module (verifiable by grep).
6. All legacy scripts live under `archive_history/`; the repo root has only the new
   framework files + the two skill JSONs + the sample config.
7. planner/schema unit tests and the designer smoke test pass.
