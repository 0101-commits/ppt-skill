"""Planning engine: client config -> normalized render spec.

Pure dict-in/dict-out. Imports NO pptx module so it can be unit-tested without
PowerPoint. Declarative (D2): every slide must already carry an explicit `type`;
the planner validates, builds the `meta` block from `identity`, auto-numbers
`section` slides, and passes slide content through untouched.
"""

from core import schema

ROMAN = ["Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ"]

# Content slide types that should carry a conclusion-sentence title (planning
# 요구사항 6 / principle P1). Sourced from the schema registry.
_TITLED_TYPES = schema.TITLED_TYPES


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
            "theme": ident.get("theme", "hcg"),      # retained for back-compat; engine ignores
            "template": ident.get("template"),       # v2.0: optional/ignored (blank 16:9 canvas)
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
