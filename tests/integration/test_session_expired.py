"""Verify SessionExpired propagates through the CLI entry point with exit 77."""

from __future__ import annotations

from unittest.mock import patch

import pytest
from typer.testing import CliRunner

from outlook_cli.cli import app, main
from outlook_cli.errors import SessionExpired


def test_session_expired_surfaces_in_clirunner() -> None:
    """Via CliRunner the exception isn't intercepted (no main() wrapper)."""
    runner = CliRunner()
    with patch(
        "outlook_cli.commands.mail.GraphClient",
        side_effect=SessionExpired("Session expired."),
    ):
        result = runner.invoke(app, ["mail", "list"])
    assert result.exception is not None
    assert isinstance(result.exception, SessionExpired)


def test_main_converts_session_expired_to_exit_77(monkeypatch: pytest.MonkeyPatch) -> None:
    """The main() wrapper installed in pyproject.toml [project.scripts] must exit 77."""

    def _raise(*_args: object, **_kwargs: object) -> None:
        raise SessionExpired("Session expired.")

    monkeypatch.setattr("outlook_cli.cli.app", _raise)
    with pytest.raises(SystemExit) as exc_info:
        main()
    assert exc_info.value.code == 77
