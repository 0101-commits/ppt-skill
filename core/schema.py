"""Slide-type registry + config/spec validation for the ppt-skill framework.

SLIDE_TYPES is the single source of truth for renderable slide types and the
keys each one needs. It mirrors `skill_ppt_design.json` -> `slide_types`, which
encodes HCG-Slide-Design-System v1.0 (the absolute standard).

Every CONTENT slide inherits the rigid skeleton, so each content type accepts
the optional skeleton fields `kicker`, `chapter`, `source` on top of its own
keys. `section` lists `num` as OPTIONAL because the planner auto-assigns Roman
numerals.

v2.0 reset: the old 19-type 4:3 set (approach_vs / diff_matrix / body_2col /
overview_3col / pain_point_categorized / container / demo_advanced ...) is
replaced by the 13 types below.
"""

# Skeleton fields every content slide may carry (engine fills the fixed frame).
_SKELETON = ["kicker", "chapter", "source"]


def _content(required, optional=None):
    """Build a content-type entry that also accepts the skeleton fields."""
    return {"required": required, "optional": (optional or []) + _SKELETON}


SLIDE_TYPES = {
    "cover":   {"required": ["title"],
                "optional": ["subtitle", "date", "confidential", "legal", "hero"]},
    "toc":     _content(["items"], ["current"]),
    "section": {"required": ["title"], "optional": ["num", "sub"]},
    "bullets": _content(["title", "items"]),
    "cards":   _content(["title", "items"]),
    "columns": _content(["title", "columns"]),
    "compare": _content(["title", "left", "right"],
                        ["left_header", "right_header", "takeaway"]),
    "kpi":     _content(["title", "items"]),
    "process": _content(["title", "steps"], ["desc"]),
    "matrix":  _content(["title", "columns", "rows"], ["takeaway"]),
    "chart":   _content(["title", "chart"]),
    "table":   _content(["title", "headers", "rows"], ["emphasis_row"]),
    "end":     {"required": [], "optional": ["message"]},
}

# Content slide types that should carry a conclusion-sentence title
# (planning 요구사항 6 / principle P1: "인사이트가 없으면 슬라이드가 아니다").
TITLED_TYPES = {
    "bullets", "cards", "columns", "compare", "kpi",
    "process", "matrix", "chart", "table",
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
    """Validate a raw client config (identity + content). Returns True or raises.

    v2.0: the renderer builds a blank 16:9 canvas from scratch, so `template` is
    no longer required. Only `identity` and `content.slides` are mandatory.
    """
    if not isinstance(cfg, dict):
        raise ConfigError("config must be a JSON object")
    for key in ("identity", "content"):
        if key not in cfg:
            raise MissingField(f"config missing top-level '{key}'")
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
