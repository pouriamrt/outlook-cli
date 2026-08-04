"""Integration tests for `outlook mail thread`."""

from __future__ import annotations

from datetime import UTC, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
from outlook_cli.graph.models import Message, Recipient
from outlook_cli.index_cache import store as store_index

runner = CliRunner()


def _msg(idx: int, text: str) -> Message:
    return Message(
        id=f"M{idx}",
        subject="Re",
        preview=text,
        importance="normal",
        is_read=True,
        is_flagged=False,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10 + idx, 0, 0, tzinfo=UTC),
        conversation_id="CV",
        **{"from": Recipient(name=f"S{idx}", address=f"s{idx}@b.c")},
        to=[],
        body_html=text,
        body_content_type="text",
    )


def test_thread_renders_all_messages_in_order(tmp_config_home: Path, tmp_cache_home: Path) -> None:
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
    store_index("mail", [{"index": 1, "id": "M-mid"}])
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.read_message", return_value=_msg(1, "first")),
        patch(
            "outlook_cli.commands.mail.read_thread",
            return_value=[_msg(1, "first"), _msg(2, "second")],
        ),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "thread", "1"])
    assert result.exit_code == 0
    assert "first" in result.stdout
    assert "second" in result.stdout
    assert result.stdout.index("first") < result.stdout.index("second")
