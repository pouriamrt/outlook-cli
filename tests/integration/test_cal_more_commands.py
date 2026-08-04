"""Integration tests for cal tomorrow / list / show / create / respond / cancel / find-time."""

from __future__ import annotations

import json
from datetime import UTC, datetime, timedelta
from pathlib import Path
from unittest.mock import MagicMock, patch

from typer.testing import CliRunner

from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app
from outlook_cli.graph.models import Event, FindTimeSuggestion, Recipient
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
    store_index("cal", [{"index": 1, "id": "EV-1"}])


def _ev() -> Event:
    return Event(
        id="EV-1",
        subject="Sync",
        start=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, 10, 30, 0, tzinfo=UTC),
        organizer=Recipient(name="P", address="p@b"),
    )


def test_cal_tomorrow(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.list_events", return_value=[_ev()]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "tomorrow"])
    assert result.exit_code == 0


def test_cal_list_range(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.list_events", return_value=[_ev()]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "list", "--start", "2026-05-22", "--end", "2026-05-23"])
    assert result.exit_code == 0


def test_cal_show(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.get_event", return_value=_ev()),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "show", "1"])
    assert result.exit_code == 0


def test_cal_show_ics(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.get_event", return_value=_ev()),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "show", "1", "--ics"])
    assert result.exit_code == 0
    assert "BEGIN:VCALENDAR" in result.stdout


def test_cal_show_json(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.get_event", return_value=_ev()),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["--json", "cal", "show", "1"])
    assert result.exit_code == 0
    assert "EV-1" in result.stdout


def test_cal_create(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.create_event", return_value="EV-NEW"),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            [
                "cal",
                "create",
                "--title",
                "Sync",
                "--start",
                "2026-05-22T10:00:00Z",
                "--duration",
                "30m",
                "--body",
                "agenda",
            ],
        )
    assert result.exit_code == 0


def test_cal_create_with_end(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.create_event", return_value="EV-NEW"),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            [
                "cal",
                "create",
                "--title",
                "Sync",
                "--start",
                "2026-05-22T10:00:00Z",
                "--end",
                "2026-05-22T11:00:00Z",
                "--body",
                "agenda",
            ],
        )
    assert result.exit_code == 0


def test_cal_create_with_invitees(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.create_event", return_value="EV-NEW") as mock_create,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            [
                "cal",
                "create",
                "--title",
                "Sync",
                "--start",
                "2026-05-22T10:00:00Z",
                "--duration",
                "30m",
                "--invitees",
                "a@b.ca,c@d.ca",
                "--body",
                "",
            ],
        )
    assert result.exit_code == 0
    spec = mock_create.call_args.args[1]
    assert "a@b.ca" in spec.invitees
    assert "c@d.ca" in spec.invitees


def test_cal_respond_accept(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.respond_event") as mock_resp,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "respond", "1", "--accept"])
    assert result.exit_code == 0
    assert mock_resp.call_args.kwargs["verb"] == "accept"


def test_cal_respond_decline(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.respond_event") as mock_resp,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "respond", "1", "--decline"])
    assert result.exit_code == 0
    assert mock_resp.call_args.kwargs["verb"] == "decline"


def test_cal_respond_tentative(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.respond_event") as mock_resp,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "respond", "1", "--tentative"])
    assert result.exit_code == 0
    assert mock_resp.call_args.kwargs["verb"] == "tentative"


def test_cal_respond_no_verb_exits_2(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["cal", "respond", "1"])
    assert result.exit_code == 2


def test_cal_respond_two_verbs_exits_2(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["cal", "respond", "1", "--accept", "--decline"])
    assert result.exit_code == 2


def test_cal_cancel(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.cancel_event") as mock_cancel,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "cancel", "1", "--comment", "no"])
    assert result.exit_code == 0
    mock_cancel.assert_called_once()


def test_cal_find_time(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.find_meeting_times", return_value=[]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(app, ["cal", "find-time", "--with", "a@b.ca", "--duration", "30m"])
    assert result.exit_code == 0


def test_cal_find_time_next_week(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.find_meeting_times", return_value=[]),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            ["cal", "find-time", "--with", "a@b.ca", "--duration", "1h", "--window", "next week"],
        )
    assert result.exit_code == 0


def test_cal_find_time_without_with_exits_2(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["cal", "find-time", "--duration", "30m"])
    assert result.exit_code == 2


def test_cal_find_time_negative_per_day_exits_2(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _setup(tmp_config_home, tmp_cache_home)
    result = runner.invoke(app, ["cal", "find-time", "--with", "a@b.ca", "--per-day", "-1"])
    assert result.exit_code == 2


def test_cal_find_time_next_week_queries_each_business_day(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    """next week → one findMeetingTimes call per Mon-Fri, first window on a Monday."""
    _setup(tmp_config_home, tmp_cache_home)
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.find_meeting_times", return_value=[]) as mock_ft,
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app, ["cal", "find-time", "--with", "a@b.ca", "--window", "next week"]
        )
    assert result.exit_code == 0
    assert mock_ft.call_count == 5  # Mon-Fri, not a single call
    first_start = mock_ft.call_args_list[0].kwargs["window_start"]
    assert first_start.weekday() == 0  # Monday alignment (was today+7)
    # Fetch the whole day's free slots (then sample per-day), not just 3 earliest.
    assert mock_ft.call_args_list[0].kwargs["max_candidates"] >= 24


def test_cal_find_time_spreads_across_the_day(tmp_config_home: Path, tmp_cache_home: Path) -> None:
    """--per-day must sample across each day (incl. afternoon), not take the earliest run."""
    _setup(tmp_config_home, tmp_cache_home)
    base = datetime(2026, 6, 1, 9, 0, tzinfo=UTC)
    # 8 consecutive 30-min free slots: 09:00 .. 12:30.
    slots = [
        FindTimeSuggestion(
            start=base + timedelta(minutes=30 * i),
            end=base + timedelta(minutes=30 * i + 30),
            confidence=100.0,
        )
        for i in range(8)
    ]
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.find_meeting_times", return_value=slots),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app,
            [
                "--json",
                "cal",
                "find-time",
                "--with",
                "a@b.ca",
                "--window",
                "next week",
                "--per-day",
                "3",
            ],
        )
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert len(data["suggestions"]) == 15  # 5 weekdays x 3
    starts = [s["start"] for s in data["suggestions"]]
    # Earliest-3 would be 09:00/09:30/10:00. Spread keeps the last slot too (12:30).
    assert any("T12:30" in s for s in starts)
    assert not any("T09:30" in s for s in starts)


def test_cal_find_time_default_shows_all_free_slots(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    """With no --per-day, every free slot is shown (no sampling) — completeness by default."""
    _setup(tmp_config_home, tmp_cache_home)
    base = datetime(2026, 6, 1, 9, 0, tzinfo=UTC)
    slots = [
        FindTimeSuggestion(
            start=base + timedelta(minutes=30 * i),
            end=base + timedelta(minutes=30 * i + 30),
            confidence=100.0,
        )
        for i in range(8)
    ]
    with (
        patch("outlook_cli.commands.calendar.GraphClient") as mock_cls,
        patch("outlook_cli.commands.calendar.find_meeting_times", return_value=slots),
    ):
        mock_cls.return_value = MagicMock()
        result = runner.invoke(
            app, ["--json", "cal", "find-time", "--with", "a@b.ca", "--window", "next week"]
        )
    assert result.exit_code == 0
    data = json.loads(result.stdout)
    assert len(data["suggestions"]) == 40  # 5 weekdays x 8, nothing dropped
