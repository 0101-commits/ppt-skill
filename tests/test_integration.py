import json
from pathlib import Path
import pytest
import main as cli

ROOT = Path(__file__).resolve().parent.parent
SAMPLE = ROOT / "projects" / "lotte_chemical.json"


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
    assert types == ["cover", "toc", "compare", "chart", "matrix",
                     "process", "kpi", "cards", "end"]


def test_full_render_pptx(tmp_path):
    # v2.0 builds a blank 16:9 canvas from scratch — no template required.
    out = tmp_path / "deck.pptx"
    rc = cli.main(["--client", "lotte_chemical", "--pptx", "--out", str(out)])
    assert rc == 0
    assert out.exists()
    from pptx import Presentation
    from pptx.util import Emu
    prs = Presentation(str(out))
    assert len(prs.slides) == 9
    # 16:9 canvas (1280x720 px = 13.333 x 7.5 in)
    assert round(prs.slide_width / 914400, 2) == 13.33
    assert round(prs.slide_height / 914400, 2) == 7.5


def test_full_render_html(tmp_path):
    out = tmp_path / "deck"
    rc = cli.main(["--client", "lotte_chemical", "--html", "--out", str(out)])
    assert rc == 0
    html = (tmp_path / "deck.html")
    assert html.exists()
    text = html.read_text(encoding="utf-8")
    assert "1280px" in text and "Pretendard" in text
    assert text.count('class="slide') == 9
