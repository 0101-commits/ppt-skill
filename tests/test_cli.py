import json
import io
import sys
from pathlib import Path
import pytest

import main as cli

ROOT = Path(__file__).resolve().parent.parent


def _write_cfg(tmp_path, slides):
    cfgdir = tmp_path / "config"
    cfgdir.mkdir()
    cfg = {"client": "demo", "identity": {"theme": "hcg", "template": "t.pptx"},
           "content": {"slides": slides}}
    (cfgdir / "demo.json").write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    return cfgdir


def test_missing_client_returns_error(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    (tmp_path / "config").mkdir()
    rc = cli.main(["--client", "ghost"])
    assert rc == 2
    assert "not found" in capsys.readouterr().err


def test_list(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "end"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    rc = cli.main(["--list"])
    assert rc == 0
    assert "demo" in capsys.readouterr().out


def test_validate_ok(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "cover", "title": "T"}, {"type": "end"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--validate"])
    assert rc == 0
    assert "ok" in capsys.readouterr().out.lower()


def test_dry_run_prints_spec(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "cover", "title": "T"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--dry-run"])
    assert rc == 0
    out = json.loads(capsys.readouterr().out)
    assert out["meta"]["theme"] == "hcg"
    assert out["slides"][0]["type"] == "cover"


def test_unknown_type_returns_config_error(monkeypatch, tmp_path, capsys):
    _write_cfg(tmp_path, [{"type": "zzz"}])
    monkeypatch.setattr(cli, "CONFIG_DIR", tmp_path / "config")
    monkeypatch.setattr(cli, "PLANNING_JSON", tmp_path / "nope.json")
    rc = cli.main(["--client", "demo", "--validate"])
    assert rc == 2
    assert "config error" in capsys.readouterr().err.lower()
