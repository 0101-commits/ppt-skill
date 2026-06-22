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
                               "optional": ["y", "h", "numbered", "bar_label", "quote", "example"]},
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
