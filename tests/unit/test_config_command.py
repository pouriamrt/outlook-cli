"""Tests for `outlook config` get / set / list commands."""

from pathlib import Path

from typer.testing import CliRunner

from outlook_cli.cli import app

runner = CliRunner()


def test_config_set_then_get(tmp_config_home: Path) -> None:
    r1 = runner.invoke(app, ["config", "set", "default_folder", "archive"])
    assert r1.exit_code == 0
    r2 = runner.invoke(app, ["config", "get", "default_folder"])
    assert r2.exit_code == 0
    assert "archive" in r2.stdout


def test_config_list_shows_defaults(tmp_config_home: Path) -> None:
    result = runner.invoke(app, ["config", "list"])
    assert result.exit_code == 0
    assert "default_folder" in result.stdout


def test_config_get_unknown_key_exits_64(tmp_config_home: Path) -> None:
    result = runner.invoke(app, ["config", "get", "nonexistent_key"])
    assert result.exit_code == 64
