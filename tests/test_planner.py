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
