"""Integration tests for `outlook cal today / tomorrow / week / list`."""

from __future__ import annotations

import json
from datetime import UTC, datetime
from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
from outlook_cli.graph.models import Event, Recipient

runner = CliRunner()


def _stub() -> None:
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


def _ev(subj: str, hour: int) -> Event:
    return Event(
        id=f"EV-{subj}",
        subject=subj,
        start=datetime(2026, 5, 22, hour, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, hour, 30, 0, tzinfo=UTC),
        organizer=Recipient(name="P", address="p@b"),
    )


def test_cal_today_uses_today_range(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub()
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch(
            "outlook_cli.commands.calendar.list_events",
            return_value=[_ev("Sync", 10), _ev("1:1", 14)],
        ),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "today"])
    assert result.exit_code == 0
    assert "Sync" in result.stdout
    assert "1:1" in result.stdout


def test_cal_today_json(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub()
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.list_events", return_value=[_ev("Sync", 10)]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["--json", "cal", "today"])
    payload = json.loads(result.stdout)
    assert payload["items"][0]["subject"] == "Sync"


def test_cal_week_uses_seven_day_range(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _stub()
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.list_events", return_value=[]) as mock_list,
    ):
        mock_cls.return_value = MagicMock()
        runner.invoke(app, ["cal", "week"])
    kwargs = mock_list.call_args.kwargs
    assert (kwargs["end"] - kwargs["start"]).days == 7
