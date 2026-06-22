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
