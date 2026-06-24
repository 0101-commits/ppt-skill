import importlib


def test_designer_imports_and_exposes_api():
    d = importlib.import_module("core.designer")
    assert d.Designer is d.DeckEngine
    for name in ("DeckEngine", "Designer", "apply_design_tokens", "apply_theme",
                 "PX", "draw_skeleton", "build_cover", "build_chart",
                 "build_compare", "build_kpi", "PALETTE", "TYPO"):
        assert hasattr(d, name), f"missing symbol {name}"


def test_legacy_entrypoints_removed():
    d = importlib.import_module("core.designer")
    assert not hasattr(d, "build"), "legacy build() must be removed"
    assert not hasattr(d, "render_spec_file"), "render_spec_file moved to main.py"
    # legacy 4:3 builders are gone
    assert not hasattr(d, "build_body_2col")
    assert not hasattr(d, "create_toc_slide")


def test_canvas_is_16_9():
    d = importlib.import_module("core.designer")
    eng = d.DeckEngine()
    assert round(eng.prs.slide_width / 914400, 2) == 13.33
    assert round(eng.prs.slide_height / 914400, 2) == 7.5


def test_apply_design_tokens_overrides_palette():
    d = importlib.import_module("core.designer")
    d.apply_design_tokens({"colors": {"primary": "#123456"}})
    assert d.PALETTE["primary"].upper().endswith("123456")
    # client override wins
    d.apply_design_tokens({"colors": {"primary": "#000000"}},
                          overrides={"primary": "#ABCDEF"})
    assert d.PALETTE["primary"].upper().endswith("ABCDEF")
    # restore standard so other tests see the real palette
    d.apply_design_tokens({"colors": {"primary": "#356CB5"}})
