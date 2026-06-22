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
