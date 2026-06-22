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
