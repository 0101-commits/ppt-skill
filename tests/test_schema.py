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
