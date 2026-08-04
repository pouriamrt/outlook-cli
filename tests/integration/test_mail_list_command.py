"""Integration tests for `outlook mail list`."""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
from outlook_cli.graph.mail import MailListResult
from outlook_cli.graph.models import Message, Recipient

runner = CliRunner()


def _stub_creds() -> None:
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


def _msg(idx: int) -> Message:
    return Message(
        id=f"ID-{idx}",
        subject=f"Subject {idx}",
        preview="p",
        importance="normal",
        is_read=False,
        is_flagged=False,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        conversation_id=f"CV-{idx}",
        **{"from": Recipient(name="A", address="a@b.ca")},
        to=[Recipient(name="P", address="pouriamortezaagha7@gmail.com")],
    )


def test_mail_list_unread_renders_table(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_creds()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_client_cls,
        patch("outlook_cli.commands.mail.list_messages") as mock_list,
        patch("outlook_cli.commands.mail.resolve_folder_id", return_value="inbox"),
    ):
        mock_client_cls.return_value = MagicMock()
        mock_list.return_value = MailListResult(items=[_msg(1), _msg(2)], next_link=None)
        result = runner.invoke(app, ["mail", "list", "--unread"])
    assert result.exit_code == 0
    assert "Subject 1" in result.stdout
    assert "Subject 2" in result.stdout


def test_mail_list_json_emits_schema(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_creds()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_client_cls,
        patch("outlook_cli.commands.mail.list_messages") as mock_list,
        patch("outlook_cli.commands.mail.resolve_folder_id", return_value="inbox"),
    ):
        mock_client_cls.return_value = MagicMock()
        mock_list.return_value = MailListResult(items=[_msg(1)], next_link=None)
        result = runner.invoke(app, ["--json", "mail", "list"])
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["items"][0]["subject"] == "Subject 1"


def test_mail_list_stores_index_cache(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub_creds()
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_client_cls,
        patch("outlook_cli.commands.mail.list_messages") as mock_list,
        patch("outlook_cli.commands.mail.resolve_folder_id", return_value="inbox"),
    ):
        mock_client_cls.return_value = MagicMock()
        mock_list.return_value = MailListResult(items=[_msg(1), _msg(2)], next_link=None)
        result = runner.invoke(app, ["mail", "list"])
    assert result.exit_code == 0
    from outlook_cli.index_cache import resolve

    assert resolve("mail", 1) == "ID-1"
    assert resolve("mail", 2) == "ID-2"
