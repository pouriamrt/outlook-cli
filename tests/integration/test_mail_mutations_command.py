"""Integration tests for mail mutation/search/reply/forward subcommands."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
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
    store_index("mail", [{"index": 1, "id": "MSG-1"}])


def test_mail_move_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.move_message") as mock_move,
        patch("outlook_cli.commands.mail.resolve_folder_id", return_value="archive"),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "move", "1", "archive"])
    assert result.exit_code == 0
    mock_move.assert_called_once()


def test_mail_delete_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.delete_message") as mock_del,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "delete", "1"])
    assert result.exit_code == 0
    mock_del.assert_called_once()


def test_mail_delete_purge_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.delete_message") as mock_del,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "delete", "1", "--purge"])
    assert result.exit_code == 0
    assert mock_del.call_args.kwargs["purge"] is True


def test_mail_flag_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.flag_message") as mock_flag,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "flag", "1"])
    assert result.exit_code == 0
    mock_flag.assert_called_once()


def test_mail_unflag_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.unflag_message") as mock_unflag,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "unflag", "1"])
    assert result.exit_code == 0
    mock_unflag.assert_called_once()


def test_mail_mark_read_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.mark_message") as mock_mark,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "mark", "1", "--read"])
    assert result.exit_code == 0
    assert mock_mark.call_args.kwargs["read"] is True


def test_mail_mark_unread_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.mark_message") as mock_mark,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "mark", "1", "--unread"])
    assert result.exit_code == 0
    assert mock_mark.call_args.kwargs["read"] is False


def test_mail_mark_without_flag_exits_2(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["mail", "mark", "1"])
    assert result.exit_code == 2


def test_mail_search_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.search_messages", return_value=[]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "search", "test"])
    assert result.exit_code == 0


def test_mail_search_with_folder(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.resolve_folder_id", return_value="archive"),
        patch("outlook_cli.commands.mail.search_messages", return_value=[]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "search", "test", "--folder", "archive"])
    assert result.exit_code == 0


def test_mail_reply_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.reply_mail") as mock_reply,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "reply", "1", "--body", "ok"])
    assert result.exit_code == 0
    mock_reply.assert_called_once()


def test_mail_reply_all_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.reply_mail") as mock_reply,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["mail", "reply", "1", "--body", "ok", "--all"])
    assert result.exit_code == 0
    assert mock_reply.call_args.kwargs["reply_all"] is True


def test_mail_forward_command(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.mail.GraphClient") as mock_cls,
        patch("outlook_cli.commands.mail.forward_mail") as mock_fwd,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app, ["mail", "forward", "1", "--to", "bob@example.com", "--body", "fyi"]
        )
    assert result.exit_code == 0
    mock_fwd.assert_called_once()


def test_mail_forward_without_to_exits_2(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["mail", "forward", "1"])
    assert result.exit_code == 2
