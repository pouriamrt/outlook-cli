"""Integration tests for `outlook mail send`."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app

runner = CliRunner()


def _stub_login() -> None:
    save(
        Credentials(
            version=1,
            acquired_at="2026-05-22T08:00:00Z",
            tenant_id="t",
            client_id="c",
            home_account_id="a.t",
            username="pouriamortezaagha7@gmail.com",
            refresh_token="rt",
        )
    )


def test_send_with_all_flags_inline(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_login()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.send_mail") as mock_send,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            [
                "mail",
                "send",
                "--to",
                "alice@example.com",
                "--subject",
                "Hello",
                "--body",
                "Plain body",
                "--importance",
                "high",
            ],
        )
    assert result.exit_code == 0
    draft = mock_send.call_args.args[1]
    assert draft.to == ["alice@example.com"]
    assert draft.subject == "Hello"
    assert draft.importance == "high"


def test_send_reads_body_from_stdin_when_dash(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_login()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.send_mail") as mock_send,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            ["mail", "send", "--to", "a@b", "--subject", "S", "--body", "-"],
            input="Body from stdin",
        )
    assert result.exit_code == 0
    assert mock_send.call_args.args[1].body == "Body from stdin"


def test_send_opens_editor_when_body_missing(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_login()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.send_mail") as mock_send,
        patch("outlook_cli.commands.mail.click.edit", return_value="From editor"),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "send", "--to", "a@b", "--subject", "S"])
    assert result.exit_code == 0
    assert mock_send.call_args.args[1].body == "From editor"


def test_send_aborts_when_body_empty_and_no_confirm(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _stub_login()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.click.edit", return_value=""),
        patch("outlook_cli.commands.mail.send_mail") as mock_send,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "send", "--to", "a@b", "--subject", "S"], input="n\n")
    assert result.exit_code == 1
    mock_send.assert_not_called()
