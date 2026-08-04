"""Integration tests for `outlook mail read`."""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
from outlook_cli.graph.models import Message, Recipient
from outlook_cli.index_cache import store as store_index

runner = CliRunner()


def _setup(tmp_config_home: Path, tmp_cache_home: Path) -> None:
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
    store_index("mail", [{"index": 1, "id": "MSG-123"}, {"index": 2, "id": "MSG-456"}])


def _msg() -> Message:
    return Message(
        id="MSG-123",
        subject="Hello",
        preview="hi",
        importance="normal",
        is_read=False,
        is_flagged=False,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        conversation_id="CV",
        **{"from": Recipient(name="Alice", address="alice@example.com")},
        to=[Recipient(name="P", address="pouriamortezaagha7@gmail.com")],
        body_html="<p>Hello!</p>",
        body_content_type="html",
    )


def test_read_by_index_resolves_via_cache(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.read_message", return_value=_msg()) as mock_read,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "read", "1"])
    assert result.exit_code == 0
    assert "Hello!" in result.stdout
    mock_read.assert_called_once()
    assert mock_read.call_args.kwargs["message_id"] == "MSG-123"


def test_read_by_raw_graph_id(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    long_id = "AAMkAGI2NzNkY2I5LWFiY2QtMTIzNA=="
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.read_message", return_value=_msg()) as mock_read,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "read", long_id])
    assert result.exit_code == 0
    assert mock_read.call_args.kwargs["message_id"] == long_id


def test_read_missing_index_exits_64(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["mail", "read", "99"])
    assert result.exit_code == 64


def test_read_json_output(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.read_message", return_value=_msg()),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["--json", "mail", "read", "1"])
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["subject"] == "Hello"
    assert "body_html" in payload or "body" in payload
