# ppt-skill Unified Framework Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Refactor the ppt-skill repo into a data-driven framework that splits Planning (config → spec) from Design (spec → pptx), preserving the proven `auto_ppt.py` rendering engine verbatim.

**Architecture:** Three core modules under `core/` — `schema.py` (slide-type registry + validation), `planner.py` (pure dict→dict structuring, no pptx), `designer.py` (the preserved engine = old `auto_ppt.py`). A `main.py` CLI loads `config/<client>.json`, runs the planner, then the designer. Legacy per-client scripts move to `archive_history/`.

**Tech Stack:** Python 3.14, python-pptx, argparse, json, pytest (tests).

## Global Constraints

- **Platform:** Windows. Use `pathlib.Path`; never hardcode `/tmp`. JSON files are UTF-8 — always open with `encoding="utf-8"`.
- **Korean text preserved:** font names (`맑은 고딕`), Roman numerals (Ⅰ Ⅱ Ⅲ Ⅳ), and all content strings stay verbatim.
- **No engine rewrite (D1):** `core/designer.py` is `auto_ppt.py` moved + minimally trimmed. Do NOT re-derive coordinates, re-author builders, or change rendering behavior. Only additive, default-safe changes allowed.
- **Declarative only (D2):** every slide carries an explicit `type`. Missing/unknown type → error. No layout inference.
- **Framework only (D3):** migrate ONLY `config/lotte_chemical.json` (representative sample). Do NOT reverse-engineer kia/paradise/lotte_aluminium decks.
- **Coords stay in Python (D4):** design JSON supplies colors/fonts at runtime; coordinate constants remain Python in `designer.py`.
- **No new heavy dependencies.** python-pptx + stdlib + pytest only.
- **pytest** run from repo root: `python -m pytest`. If missing: `pip install pytest`.
- **Spec:** `docs/superpowers/specs/2026-06-19-ppt-skill-refactor-design.md` (authority for acceptance criteria §12).
- **Branch:** all work on `refactor/unified-framework`.

---

### Task 1: Scaffold directories + archive legacy files

**Files:**
- Create: `core/__init__.py`, `config/.gitkeep`, `archive_history/README.md`
- Create: `archive_history/auto_ppt_legacy.py` (snapshot copy of `auto_ppt.py`)
- Move (git mv): legacy scripts + analysis artifacts per spec §4.1

**Interfaces:**
- Produces: the `core/` package directory (importable), an `archive_history/` holding frozen scripts, `auto_ppt.py` still present at root (Task 2 transforms it).

- [ ] **Step 1: Create package + dirs**

```bash
cd /c/Users/cgpar/ppt-skill
mkdir -p core config archive_history
printf '"""ppt-skill core package: planner, designer, schema."""\n' > core/__init__.py
> config/.gitkeep
```

- [ ] **Step 1b: Fix .gitignore so package files are trackable**

The existing `.gitignore` has a `_*.py` scratch rule that also matches `__init__.py`
(it starts with `_`), which would make `git add core/__init__.py` a silent no-op.
Edit `.gitignore`: change the line `_*.py` to `_[a-z]*.py` (still ignores scratch
files like `_foo.py`, no longer ignores dunder files). Leave the rest untouched
(`final_analysis.json` stays ignored — it is a regenerable dump).

- [ ] **Step 2: Snapshot the engine before any edits (safety copy)**

```bash
cp auto_ppt.py archive_history/auto_ppt_legacy.py
```

- [ ] **Step 3: Move legacy scripts + analysis artifacts (history preserved)**

```bash
cd /c/Users/cgpar/ppt-skill
git mv auto_ppt_kia.py lotte_chemical_ppt.py paradise_compare.py archive_history/
git mv analyze_final.py report2.py report_final.py archive_history/
git mv spec_paradise.json paradise_compare.json archive_history/
git mv STRUCTURE_REPORT.md STRUCTURE_REPORT2.md PARADISE_COMPARE.md archive_history/
# final_analysis.json is gitignored (untracked) — plain mv; it stays ignored in its new home
mv final_analysis.json archive_history/ 2>/dev/null || true
git add archive_history/auto_ppt_legacy.py archive_history/README.md core/__init__.py config/.gitkeep .gitignore
```

Note: `skill_ppt_planning.json`, `skill_ppt_design.json`, their `.md` pairs, `README.md`, `auto_ppt.py`, and the Draft `.pptx` files stay at root (per §4.1). If any tracked file above does not exist, drop it from the command — do not fail the task. (`git mv` only works on git-tracked files; `final_analysis.json` is handled separately above because it is gitignored.)

- [ ] **Step 4: Write the archive README**

```markdown
# archive_history

Pre-refactor scripts and analysis artifacts. **Frozen** — these target the old
`auto_ppt` module API (before the 2026-06-19 unified-framework refactor) and are
kept for reference and future migration to `config/*.json`.

| File | What it was |
|------|-------------|
| `auto_ppt_legacy.py` | Snapshot of the monolith `auto_ppt.py` before it became `core/designer.py` |
| `auto_ppt_kia.py` | 기아 imperative deck generator |
| `lotte_chemical_ppt.py` | 롯데케미칼 imperative deck generator |
| `paradise_compare.py` | 파라다이스 Draft-vs-final QA/analysis tool |
| `analyze_final.py`, `report2.py`, `report_final.py` | one-off structure-analysis scripts |
| `final_analysis.json`, `spec_paradise.json`, `paradise_compare.json` | analysis/paradise data |
| `STRUCTURE_REPORT*.md`, `PARADISE_COMPARE.md` | analysis reports |

To revive a client: author `config/<client>.json` and render via `python main.py --client <client>`.
```

Write that content to `archive_history/README.md`.

- [ ] **Step 5: Verify tree, then commit**

```bash
cd /c/Users/cgpar/ppt-skill
ls core config archive_history
git status --short
git commit -m "refactor: 디렉터리 스캐폴딩 + 레거시 스크립트 archive_history 이관

core/ config/ archive_history/ 생성. 레거시 클라이언트/분석 스크립트 및 산출물 이관(§4.1).
auto_ppt.py 스냅샷을 archive_history/auto_ppt_legacy.py로 보존.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

Expected: `core/`, `config/`, `archive_history/` exist; archived files listed under `archive_history/`; `auto_ppt.py` still at root.

---

### Task 2: Transform auto_ppt.py → core/designer.py (preserve engine)

**Files:**
- Move: `auto_ppt.py` → `core/designer.py`
- Modify: `core/designer.py` (trim legacy entrypoints; add `Designer` alias + `apply_design_tokens` + `design_tokens` param)
- Test: `tests/test_designer_smoke.py`

**Interfaces:**
- Consumes: nothing from other tasks.
- Produces:
  - `core.designer.DeckEngine(template, out, theme="hcg", design_tokens=None, color_overrides=None)` with `.render(spec) -> self` and `.save(out=None)`.
  - `core.designer.Designer` = alias of `DeckEngine`.
  - `core.designer.apply_design_tokens(tokens: dict, overrides: dict | None = None) -> None` — updates the module color palette from parsed `skill_ppt_design.json`.
  - All existing builders/helpers/`apply_theme` remain importable and unchanged.

- [ ] **Step 1: Write the failing smoke test**

`tests/test_designer_smoke.py`:

```python
import importlib

def test_designer_imports_and_exposes_api():
    d = importlib.import_module("core.designer")
    assert hasattr(d, "DeckEngine")
    assert hasattr(d, "Designer")
    assert d.Designer is d.DeckEngine
    assert hasattr(d, "apply_design_tokens")
    assert hasattr(d, "apply_theme")
    # core builders preserved
    for name in ("add_item", "build_body_2col", "build_approach_vs",
                 "set_text_outline", "add_shadow", "create_toc_slide"):
        assert hasattr(d, name), f"missing preserved symbol {name}"

def test_legacy_entrypoints_removed():
    d = importlib.import_module("core.designer")
    assert not hasattr(d, "build"), "legacy build() must be removed"
    assert not hasattr(d, "render_spec_file"), "render_spec_file moved to main.py"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_designer_smoke.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'core.designer'`.

- [ ] **Step 3: Move the engine into the package**

```bash
cd /c/Users/cgpar/ppt-skill
git mv auto_ppt.py core/designer.py
```

- [ ] **Step 4: Trim legacy entrypoints (content-anchored, not line numbers)**

In `core/designer.py`, locate by content and DELETE these only:
1. The legacy `def build(` function (the 롯데알미늄 hardcoded deck builder).
2. The `def render_spec_file(` function (CLI wrapper — `main.py` replaces it).
3. The trailing `if __name__ == "__main__":` block.
4. The module-level template-path constants `REAL = ...`, `KIA_FINAL = ...`, and `OUT = ...`.

Before deleting the constants, confirm no RETAINED code uses them:

```bash
cd /c/Users/cgpar/ppt-skill
grep -nE '\b(REAL|KIA_FINAL|OUT)\b' core/designer.py
```

If a retained builder still references one, replace that constant with `= None` instead of deleting it, and note it. Keep EVERYTHING else: all color constants (`HCG_RED`, `KIA_*`, `PARADISE_*`), coordinate constants, `THEME`/`apply_theme`, all `_*`/`add_*`/`build_*` functions, the advanced diagram builders, and the `DeckEngine` class.

- [ ] **Step 5: Add the `Designer` alias + design-token consumption (additive, default-safe)**

At the END of `core/designer.py`, append:

```python
# --- Unified-framework additions (additive; default behavior unchanged) ---

Designer = DeckEngine  # public alias for the framework


def apply_design_tokens(tokens, overrides=None):
    """Update the module color palette from parsed skill_ppt_design.json.

    `tokens` is the parsed design JSON; reads tokens["theme_colors"], where each
    value is either a hex string or a dict with an "rgb" key. `overrides` is an
    optional {NAME: "#RRGGBB"} map (e.g. client identity.colors) applied last.
    No-op for names not present. Mirrors apply_theme()'s global-rebind pattern.
    """
    global HCG_RED
    resolved = {}
    for name, val in (tokens.get("theme_colors") or {}).items():
        hexv = val.get("rgb") if isinstance(val, dict) else val
        if hexv:
            resolved[name] = str(hexv).lstrip("#")
    for name, hexv in (overrides or {}).items():
        if hexv:
            resolved[name] = str(hexv).lstrip("#")
    if "HCG_RED" in resolved:
        HCG_RED = RGBColor.from_string(resolved["HCG_RED"])
    for key, hexv in resolved.items():
        if isinstance(THEME, dict) and key in THEME:
            THEME[key] = RGBColor.from_string(hexv)
```

Note: `RGBColor` and `THEME` already exist in the module (imported/defined above). If the module imports `RGBColor` under a different name, use that name.

- [ ] **Step 6: Add `design_tokens`/`color_overrides` params to `DeckEngine.__init__`**

Find `class DeckEngine` and its `def __init__(self, template, out, theme=...):`. Change the signature to:

```python
    def __init__(self, template, out, theme="hcg", design_tokens=None, color_overrides=None):
```

Immediately AFTER the existing `apply_theme(theme)` call inside `__init__`, add:

```python
        if design_tokens or color_overrides:
            apply_design_tokens(design_tokens or {}, color_overrides)
```

Do not change anything else in `__init__`.

- [ ] **Step 7: Run the smoke test**

Run: `python -m pytest tests/test_designer_smoke.py -v`
Expected: PASS (both tests).

- [ ] **Step 8: Add a token-consumption unit test, run it**

Append to `tests/test_designer_smoke.py`:

```python
def test_apply_design_tokens_overrides_palette():
    d = importlib.import_module("core.designer")
    before = str(d.HCG_RED)
    d.apply_design_tokens({"theme_colors": {"HCG_RED": {"rgb": "#123456"}}})
    assert str(d.HCG_RED) != before
    assert str(d.HCG_RED).upper().endswith("123456")
    # client override wins
    d.apply_design_tokens({"theme_colors": {"HCG_RED": "#000000"}},
                          overrides={"HCG_RED": "#ABCDEF"})
    assert str(d.HCG_RED).upper().endswith("ABCDEF")
```

Run: `python -m pytest tests/test_designer_smoke.py -v`
Expected: PASS (3 tests).

- [ ] **Step 9: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add core/designer.py tests/test_designer_smoke.py
git commit -m "refactor: auto_ppt.py -> core/designer.py (엔진 보존) + 디자인토큰 소비

레거시 build()/render_spec_file()/__main__/템플릿경로 상수 제거.
Designer 별칭 + apply_design_tokens() 추가, DeckEngine.__init__에 design_tokens/color_overrides(기본 None) 추가. 렌더링 동작 불변.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

### Task 3: core/schema.py — slide-type registry + validation

**Files:**
- Create: `core/schema.py`
- Test: `tests/test_schema.py`

**Interfaces:**
- Produces:
  - `core.schema.SLIDE_TYPES: dict[str, dict]` — `{type: {"required": [...], "optional": [...]}}`.
  - `core.schema.validate_config(cfg: dict) -> True` (raises on invalid).
  - `core.schema.validate_spec(spec: dict) -> True` (raises on invalid).
  - Exceptions `ConfigError`, `UnknownSlideType(ConfigError)`, `MissingField(ConfigError)`.

- [ ] **Step 1: Write failing tests**

`tests/test_schema.py`:

```python
import pytest
from core import schema


def _valid_cfg():
    return {
        "client": "x",
        "identity": {"theme": "hcg", "template": "t.pptx"},
        "content": {"slides": [
            {"type": "cover", "title": "T"},
            {"type": "end"},
        ]},
    }


def test_valid_config_passes():
    assert schema.validate_config(_valid_cfg()) is True


def test_missing_identity_raises():
    cfg = _valid_cfg(); del cfg["identity"]
    with pytest.raises(schema.MissingField):
        schema.validate_config(cfg)


def test_unknown_slide_type_raises():
    cfg = _valid_cfg()
    cfg["content"]["slides"].append({"type": "nope"})
    with pytest.raises(schema.UnknownSlideType):
        schema.validate_config(cfg)


def test_missing_required_field_raises():
    cfg = _valid_cfg()
    cfg["content"]["slides"] = [{"type": "cover"}]  # 'title' missing
    with pytest.raises(schema.MissingField):
        schema.validate_config(cfg)


def test_slide_missing_type_raises():
    cfg = _valid_cfg()
    cfg["content"]["slides"] = [{"title": "no type"}]
    with pytest.raises(schema.MissingField):
        schema.validate_config(cfg)


def test_registry_has_19_types():
    assert len(schema.SLIDE_TYPES) == 19
    assert "approach_vs" in schema.SLIDE_TYPES
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_schema.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'core.schema'`.

- [ ] **Step 3: Implement core/schema.py**

```python
"""Slide-type registry + config/spec validation for the ppt-skill framework.

SLIDE_TYPES is the single source of truth for the renderable slide types and the
keys each one needs. The required/optional key sets were derived from the
DeckEngine.render() dispatch in core/designer.py. `section` lists `num` as
OPTIONAL because the planner auto-assigns Roman numerals.
"""

SLIDE_TYPES = {
    "cover":                  {"required": ["title"], "optional": ["subtitle", "date"]},
    "toc":                    {"required": ["items"], "optional": ["title", "current"]},
    "section_agenda":         {"required": ["sections", "current"], "optional": []},
    "blocks":                 {"required": ["title", "subtitle", "items"],
                               "optional": ["y", "h", "bar_label", "quote", "example"]},
    "container":              {"required": ["title", "subtitle", "containers"], "optional": []},
    "overview":               {"required": ["subtitle"],
                               "optional": ["background", "scope", "sub_details", "plan", "quote"]},
    "section":                {"required": ["title"], "optional": ["num", "sub"]},
    "body_2col":              {"required": ["title", "subtitle", "header_l", "header_r", "left", "right"],
                               "optional": []},
    "body_single":            {"required": ["title", "subtitle", "rows"], "optional": []},
    "body_process":           {"required": ["title", "subtitle", "steps"], "optional": ["desc"]},
    "overview_3col":          {"required": ["title", "subtitle", "cols"],
                               "optional": ["bar_label", "example"]},
    "diff_matrix":            {"required": ["title", "subtitle", "rows"],
                               "optional": ["quote", "example"]},
    "pain_point_categorized": {"required": ["title", "subtitle", "left_hdr", "right_hdr", "rows"],
                               "optional": ["summary_left", "summary_right"]},
    "approach_vs":            {"required": ["title", "subtitle", "left_title", "left",
                                            "right_title", "right"], "optional": ["quote"]},
    "process_roadmap":        {"required": ["title", "subtitle", "phases"], "optional": []},
    "compare_table":          {"required": ["title", "subtitle", "headers", "rows"],
                               "optional": ["example"]},
    "appendix":               {"required": ["rows"], "optional": ["title", "subtitle"]},
    "demo_advanced":          {"required": [], "optional": ["title", "subtitle"]},
    "end":                    {"required": [], "optional": []},
}


class ConfigError(Exception):
    """Base error for malformed client config or spec."""


class UnknownSlideType(ConfigError):
    """A slide declared a type not in SLIDE_TYPES."""


class MissingField(ConfigError):
    """A required key is absent from config or a slide."""


def _validate_slide(slide, idx):
    if not isinstance(slide, dict) or "type" not in slide:
        raise MissingField(f"slide[{idx}] missing 'type'")
    stype = slide["type"]
    if stype not in SLIDE_TYPES:
        raise UnknownSlideType(
            f"slide[{idx}] unknown type '{stype}'. Valid: {sorted(SLIDE_TYPES)}"
        )
    for field in SLIDE_TYPES[stype]["required"]:
        if field not in slide:
            raise MissingField(
                f"slide[{idx}] type '{stype}' missing required '{field}'"
            )


def validate_config(cfg):
    """Validate a raw client config (identity + content). Returns True or raises."""
    if not isinstance(cfg, dict):
        raise ConfigError("config must be a JSON object")
    for key in ("identity", "content"):
        if key not in cfg:
            raise MissingField(f"config missing top-level '{key}'")
    for key in ("theme", "template"):
        if key not in cfg["identity"]:
            raise MissingField(f"identity missing '{key}'")
    slides = cfg["content"].get("slides")
    if not isinstance(slides, list):
        raise MissingField("content.slides must be a list")
    for i, slide in enumerate(slides):
        _validate_slide(slide, i)
    return True


def validate_spec(spec):
    """Validate a planner-produced spec (meta + slides). Returns True or raises."""
    if not isinstance(spec, dict) or "meta" not in spec or "slides" not in spec:
        raise ConfigError("spec must have 'meta' and 'slides'")
    for i, slide in enumerate(spec["slides"]):
        _validate_slide(slide, i)
    return True
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_schema.py -v`
Expected: PASS (6 tests).

- [ ] **Step 5: Reconcile registry against the real engine**

Open `core/designer.py`, find `DeckEngine.render` and the `build_*` functions it dispatches to. For each of the 19 types, confirm the `required`/`optional` keys in `SLIDE_TYPES` match the keys the builder actually reads (`spec_slide["..."]` / `.get("...")`). Fix any mismatch in `SLIDE_TYPES`, re-run Step 4. Record any change in the commit message.

- [ ] **Step 6: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add core/schema.py tests/test_schema.py
git commit -m "feat: core/schema.py — 19개 슬라이드 타입 레지스트리 + 검증

알 수 없는 타입/누락 필드를 명시적 예외로(기존 silent skip 버그 해결).

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

### Task 4: core/planner.py — config → normalized spec (no pptx)

**Files:**
- Create: `core/planner.py`
- Test: `tests/test_planner.py`

**Interfaces:**
- Consumes: `core.schema.validate_config`, `validate_spec`, error classes.
- Produces:
  - `core.planner.Planner(planning_rules: dict | None = None)`.
  - `Planner.plan(config: dict, out_override: str | None = None) -> dict` returning `{"meta": {...}, "slides": [...]}`.
  - `Planner.warnings: list[str]` populated after `plan()`.
  - `meta` keys: `theme`, `template`, `out`, `colors` (dict), `fonts` (dict).

- [ ] **Step 1: Write failing tests**

`tests/test_planner.py`:

```python
import sys
import pytest
from core.planner import Planner
from core import schema


def _cfg(slides, **identity):
    ident = {"theme": "hcg", "template": "t.pptx"}
    ident.update(identity)
    return {"client": "acme", "identity": ident, "content": {"slides": slides}}


def test_plan_builds_meta_from_identity():
    spec = Planner().plan(_cfg([{"type": "end"}],
                               colors={"primary": "#921F0B"}, fonts={"kr": "맑은 고딕"}))
    assert spec["meta"]["theme"] == "hcg"
    assert spec["meta"]["template"] == "t.pptx"
    assert spec["meta"]["colors"] == {"primary": "#921F0B"}
    assert spec["meta"]["fonts"] == {"kr": "맑은 고딕"}


def test_out_override_wins():
    spec = Planner().plan(_cfg([{"type": "end"}]), out_override="X.pptx")
    assert spec["meta"]["out"] == "X.pptx"


def test_out_defaults_to_client_name():
    spec = Planner().plan(_cfg([{"type": "end"}]))
    assert spec["meta"]["out"].endswith("acme.pptx")


def test_sections_auto_numbered_roman():
    spec = Planner().plan(_cfg([
        {"type": "section", "title": "배경"},
        {"type": "section", "title": "방안"},
        {"type": "section", "title": "일정", "num": "Ⅸ"},  # explicit kept
    ]))
    nums = [s["num"] for s in spec["slides"] if s["type"] == "section"]
    assert nums == ["Ⅰ", "Ⅱ", "Ⅸ"]


def test_input_config_not_mutated():
    cfg = _cfg([{"type": "section", "title": "배경"}])
    Planner().plan(cfg)
    assert "num" not in cfg["content"]["slides"][0]


def test_invalid_config_propagates():
    with pytest.raises(schema.UnknownSlideType):
        Planner().plan(_cfg([{"type": "bogus"}]))


def test_planner_imports_no_pptx():
    # Importing planner alone must not pull in python-pptx.
    for m in list(sys.modules):
        if m == "pptx" or m.startswith("pptx."):
            del sys.modules[m]
    import importlib
    importlib.reload(sys.modules["core.planner"])
    assert "pptx" not in sys.modules
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_planner.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'core.planner'`.

- [ ] **Step 3: Implement core/planner.py**

```python
"""Planning engine: client config -> normalized render spec.

Pure dict-in/dict-out. Imports NO pptx module so it can be unit-tested without
PowerPoint. Declarative (D2): every slide must already carry an explicit `type`;
the planner validates, builds the `meta` block from `identity`, auto-numbers
`section` slides, and passes slide content through untouched.
"""

from core import schema

ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ"]

# Content slide types that should carry a title (planning principle P1:
# "인사이트가 없으면 슬라이드가 아니다").
_TITLED_TYPES = {
    "blocks", "body_2col", "body_single", "body_process", "overview_3col",
    "diff_matrix", "pain_point_categorized", "approach_vs", "process_roadmap",
    "compare_table",
}


class Planner:
    def __init__(self, planning_rules=None):
        self.rules = planning_rules or {}
        self.warnings = []

    def plan(self, config, out_override=None):
        schema.validate_config(config)
        self.warnings = []
        ident = config["identity"]
        content = config["content"]

        meta = {
            "theme": ident.get("theme", "hcg"),
            "template": ident["template"],
            "out": out_override or content.get("out") or f"{config.get('client', 'output')}.pptx",
            "colors": ident.get("colors", {}),
            "fonts": ident.get("fonts", {}),
        }

        slides = self._structure(content["slides"])
        spec = {"meta": meta, "slides": slides}
        schema.validate_spec(spec)
        return spec

    def _structure(self, raw_slides):
        out = []
        section_n = 0
        for raw in raw_slides:
            slide = dict(raw)  # shallow copy — never mutate caller's config
            if slide["type"] == "section":
                if not slide.get("num"):
                    slide["num"] = (ROMAN[section_n] if section_n < len(ROMAN)
                                    else str(section_n + 1))
                section_n += 1
            if slide["type"] in _TITLED_TYPES and not slide.get("title"):
                self.warnings.append(
                    f"slide type '{slide['type']}' has no title (planning P1)"
                )
            out.append(slide)
        self._check_storyline(section_n)
        return out

    def _check_storyline(self, section_count):
        """Light cross-check against skill_ppt_planning.json storyline, if present."""
        arch = self.rules.get("storyline_architecture", {})
        sections = arch.get("섹션") or arch.get("sections")
        if isinstance(sections, list) and sections:
            expected = sum(1 for s in sections
                           if str(s.get("id", "")).startswith("S"))
            if expected and section_count and section_count != expected:
                self.warnings.append(
                    f"deck has {section_count} section dividers; "
                    f"storyline defines {expected}"
                )
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_planner.py -v`
Expected: PASS (7 tests).

- [ ] **Step 5: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add core/planner.py tests/test_planner.py
git commit -m "feat: core/planner.py — 선언적 기획 엔진(config→spec, pptx 비의존)

meta 구성 + section 로마숫자 자동번호 + P1 경고. 입력 config 불변.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

### Task 5: main.py — unified CLI entry point

**Files:**
- Create: `main.py`
- Test: `tests/test_cli.py`

**Interfaces:**
- Consumes: `core.planner.Planner`, `core.schema`, `core.designer.Designer`.
- Produces: `main.main(argv: list[str] | None = None) -> int` (exit code) and a `__main__` guard. CLI flags: `--client`, `--list`, `--out`, `--dry-run`, `--validate`.

- [ ] **Step 1: Write failing tests**

`tests/test_cli.py`:

```python
import json
import io
import sys
from pathlib import Path
import pytest

import main as cli

ROOT = Path(__file__).resolve().parent.parent


def _write_cfg(tmp_path, slides):
    cfgdir = tmp_path / "config"
    cfgdir.mkdir()
    cfg = {"client": "demo", "identity": {"theme": "hcg", "template": "t.pptx"},
           "content": {"slides": slides}}
    (cfgdir / "demo.json").write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    return cfgdir


def test_missing_client_returns_error(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    (tmp_path / "config").mkdir()
    rc = cli.main(["--client", "ghost"])
    assert rc == 2
    assert "not found" in capsys.readouterr().err


def test_list(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "end"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    rc = cli.main(["--list"])
    assert rc == 0
    assert "demo" in capsys.readouterr().out


def test_validate_ok(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "cover", "title": "T"}, {"type": "end"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--validate"])
    assert rc == 0
    assert "ok" in capsys.readouterr().out.lower()


def test_dry_run_prints_spec(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "cover", "title": "T"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--dry-run"])
    assert rc == 0
    out = json.loads(capsys.readouterr().out)
    assert out["meta"]["theme"] == "hcg"
    assert out["slides"][0]["type"] == "cover"


def test_unknown_type_returns_config_error(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "zzz"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--validate"])
    assert rc == 2
    assert "config error" in capsys.readouterr().err.lower()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'main'` (or AttributeError on `CONFIG_DIR`).

- [ ] **Step 3: Implement main.py**

```python
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_cli.py -v`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add main.py tests/test_cli.py
git commit -m "feat: main.py — 통합 CLI(--client/--list/--out/--dry-run/--validate)

config 로드→planner→designer 파이프라인. 누락 config/템플릿/잘못된 타입에 친절한 에러.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

### Task 6: Sample config + end-to-end integration

**Files:**
- Create: `config/lotte_chemical.json`
- Test: `tests/test_integration.py`

**Interfaces:**
- Consumes: full pipeline (main → planner → designer).
- Produces: a working representative client config; an integration test proving validate/dry-run/list always work and full render works when the template is present.

- [ ] **Step 1: Reconcile nested slide shapes against the engine**

Open `core/designer.py` and read the builders for the 6 types used below
(`cover`, `toc`, `overview`, `approach_vs`, `diff_matrix`, `end`). Note the exact
nested shape each expects (e.g. what an `overview` `scope`/`plan` value is — list vs
string; what each `diff_matrix` `rows` element looks like — dict keys vs list). Adjust
the sample in Step 2 so nested shapes match the builders. (Top-level required keys are
already correct per `SLIDE_TYPES`.)

- [ ] **Step 2: Write the sample config**

`config/lotte_chemical.json` — adjust nested shapes per Step 1 if they differ:

```json
{
  "client": "lotte_chemical",
  "identity": {
    "display_name": "롯데케미칼",
    "theme": "hcg",
    "template": "C:\\Users\\cgpar\\OneDrive - 휴먼컨설팅그룹\\09 Admin\\09 etc\\other\\Claude\\롯데알미늄 제안서\\[HCG] 롯데알미늄_직무기반 HR제도 설계 및 도입_제안서_final.pptx",
    "colors": {"HCG_RED": "#921F0B"},
    "fonts": {"kr": "맑은 고딕", "en": "Arial"},
    "logo": ""
  },
  "content": {
    "out": "HCG_롯데케미칼_직무기반HR_제안서_Draft.pptx",
    "slides": [
      {"type": "cover",
       "title": "롯데케미칼 직무기반 HR제도 설계 및 도입",
       "subtitle": "- 제안서 -",
       "date": "2025.06"},

      {"type": "toc",
       "items": ["프로젝트 배경", "추진 방안", "일정 및 조직", "HCG 소개"]},

      {"type": "overview",
       "subtitle": "직무 중심으로 전환하는 HR 운영 체계",
       "background": "연공 중심 보상의 한계와 사업 포트폴리오 다변화",
       "scope": "직무분석 · 직무평가 · 직무급 설계 · 제도 이행",
       "plan": "3단계 · 16주 로드맵으로 현업 수용성을 확보하며 단계적 전환"},

      {"type": "approach_vs",
       "title": "왜 별도 직무체계 프로젝트인가",
       "subtitle": "기존 직급제도 보완 vs 직무기반 재설계",
       "left_title": "직급제도 부분 보완",
       "left": ["기존 호봉 구조 유지", "단기 갈등 최소화", "근본 원인 미해결"],
       "right_title": "직무기반 재설계",
       "right": ["직무가치 기반 보상", "성과·역량 연계", "지속가능한 운영체계"],
       "quote": "보상의 공정성은 직무가치에서 출발한다"},

      {"type": "diff_matrix",
       "title": "직군별 직무 특성과 적용 방향",
       "subtitle": "직군 · 특성 · 시사점 · 적용",
       "rows": [
         {"group": "연구·기술", "trait": "전문성 심화 경로", "insight": "이중 사다리 필요", "apply": "전문직 트랙 신설"},
         {"group": "생산·운영", "trait": "교대·숙련 중심", "insight": "숙련 가치 반영", "apply": "숙련급 구간 설계"},
         {"group": "경영지원", "trait": "직무 범위 광범위", "insight": "직무 표준화 우선", "apply": "직무기술서 정비"}
       ],
       "quote": "직군 특성을 무시한 일률 적용은 실패한다"},

      {"type": "end"}
    ]
  }
}
```

- [ ] **Step 3: Write the integration test**

`tests/test_integration.py`:

```python
import json
from pathlib import Path
import pytest
import main as cli

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "config" / "lotte_chemical.json"


def test_sample_lists():
    assert "lotte_chemical" in cli.list_clients()


def test_sample_validates():
    rc = cli.main(["--client", "lotte_chemical", "--validate"])
    assert rc == 0


def test_sample_dry_run_spec_shape(capsys):
    rc = cli.main(["--client", "lotte_chemical", "--dry-run"])
    assert rc == 0
    spec = json.loads(capsys.readouterr().out)
    types = [s["type"] for s in spec["slides"]]
    assert types == ["cover", "toc", "overview", "approach_vs", "diff_matrix", "end"]
    assert spec["meta"]["theme"] == "hcg"


def test_full_render_if_template_present(tmp_path):
    cfg = json.loads(SAMPLE.read_text(encoding="utf-8"))
    template = Path(cfg["identity"]["template"])
    if not template.exists():
        pytest.skip(f"template not present: {template}")
    out = tmp_path / "deck.pptx"
    rc = cli.main(["--client", "lotte_chemical", "--out", str(out)])
    assert rc == 0
    assert out.exists()
    from pptx import Presentation
    prs = Presentation(str(out))
    assert len(prs.slides) >= 6
```

- [ ] **Step 4: Run the integration test**

Run: `python -m pytest tests/test_integration.py -v`
Expected: 3 PASS; `test_full_render_if_template_present` PASSES if the template exists, otherwise SKIPS. If it FAILS on render, return to Step 1 and fix nested shapes to match the builders.

- [ ] **Step 5: Run the full suite**

Run: `python -m pytest -v`
Expected: all tests pass (renders skip only if template absent).

- [ ] **Step 6: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add config/lotte_chemical.json tests/test_integration.py
git commit -m "feat: config/lotte_chemical.json 샘플 + 엔드투엔드 통합 테스트

6개 슬라이드 타입으로 데이터 매핑 시연. validate/dry-run/list 항상 검증, 전체 렌더는 템플릿 존재 시.

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

---

### Task 7: README update + acceptance verification

**Files:**
- Modify: `README.md`
- Verify: spec §12 acceptance criteria

**Interfaces:**
- Consumes: everything. No new code.

- [ ] **Step 1: Rewrite README usage section**

Replace the usage portion of `README.md` with:

```markdown
## 사용법 (통합 프레임워크)

    python main.py --list                      # 등록된 클라이언트 목록
    python main.py --client lotte_chemical      # 렌더링
    python main.py --client lotte_chemical --dry-run    # 정규화 spec만 출력
    python main.py --client lotte_chemical --validate   # 검증만
    python main.py --client lotte_chemical --out deck.pptx

### 구조
- `main.py` — CLI 진입점
- `core/planner.py` — 기획 엔진 (config → spec, pptx 비의존)
- `core/designer.py` — 디자인 엔진 (spec → pptx, 검증된 DeckEngine)
- `core/schema.py` — 슬라이드 타입 레지스트리 + 검증
- `config/<client>.json` — 클라이언트 아이덴티티 + 콘텐츠
- `skill_ppt_planning.json` / `skill_ppt_design.json` — 기획/디자인 규칙
- `archive_history/` — 레거시 스크립트 (동결)

### 새 클라이언트 추가
`config/<client>.json` 한 개만 작성 (Python 코드 불필요). 스키마는
`docs/superpowers/specs/2026-06-19-ppt-skill-refactor-design.md` §7 참고.
```

- [ ] **Step 2: Verify acceptance criteria (spec §12)**

Run each and confirm:

```bash
cd /c/Users/cgpar/ppt-skill
python main.py --list                                   # shows: lotte_chemical
python main.py --client lotte_chemical --validate        # [ok] ... slides valid
python main.py --client lotte_chemical --dry-run         # prints spec JSON
python -m pytest -v                                      # all pass
grep -rn "import pptx\|from pptx" core/planner.py core/schema.py || echo "OK: no pptx in planner/schema"
ls archive_history                                       # legacy scripts present
ls core                                                  # __init__.py designer.py planner.py schema.py
```

Confirm against §12: (1) render produces pptx [if template present], (2) `--list` shows lotte_chemical, (3) `--dry-run` prints spec, (4) unknown type → error (tested in `test_cli.py`), (5) no pptx import in planner/schema, (6) root dispositioned per §4.1, (7) tests pass.

- [ ] **Step 3: Commit**

```bash
cd /c/Users/cgpar/ppt-skill
git add README.md
git commit -m "docs: README 통합 프레임워크 사용법 갱신

Co-Authored-By: Claude Opus 4.8 (1M context) <noreply@anthropic.com>"
```

- [ ] **Step 4: Final branch state**

```bash
cd /c/Users/cgpar/ppt-skill
git log --oneline refactor/unified-framework
git status
```

Expected: clean tree; commit per task. Branch ready for review/merge (do NOT merge or push unless asked).

---

## Notes for the executor

- **Preserve discipline:** Task 2 is a MOVE, not a rewrite. If tempted to "clean up" designer.py internals, stop — that's out of scope (D1).
- **Template dependency:** full rendering needs the real 롯데알미늄 `_final.pptx` template (private path in the sample config). Without it, validate/dry-run/list still prove the framework; the render test skips.
- **Registry truth:** `SLIDE_TYPES` (Task 3) and nested shapes (Task 6) are reconciled against `core/designer.py` source — the engine is the authority, not this plan's best-effort key lists.
- **Warnings:** planner collects `warnings` (non-fatal). Surfacing them in the CLI is optional polish, not required for acceptance.
