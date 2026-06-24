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


def test_registry_has_expected_types():
    assert len(schema.SLIDE_TYPES) == 13
    for t in ("cover", "toc", "section", "bullets", "cards", "columns",
              "compare", "kpi", "process", "matrix", "chart", "table", "end"):
        assert t in schema.SLIDE_TYPES
    # legacy 4:3 types are gone
    assert "approach_vs" not in schema.SLIDE_TYPES


def test_content_types_accept_skeleton_fields():
    for t in ("bullets", "chart", "compare"):
        opt = schema.SLIDE_TYPES[t]["optional"]
        assert {"kicker", "chapter", "source"} <= set(opt)
