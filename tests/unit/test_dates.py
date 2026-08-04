from datetime import UTC, datetime, timedelta

from freezegun import freeze_time

from outlook_cli.dates import parse_human, parse_since


@freeze_time("2026-05-22T10:00:00", tz_offset=0)
def test_parse_human_handles_iso() -> None:
    dt = parse_human("2026-05-23T15:00:00Z")
    assert dt.year == 2026 and dt.month == 5 and dt.day == 23


@freeze_time("2026-05-22T10:00:00", tz_offset=0)
def test_parse_human_handles_relative() -> None:
    dt = parse_human("tomorrow 3pm")
    assert dt.day == 23
    assert dt.hour == 15


@freeze_time("2026-05-22T10:00:00Z")
def test_parse_since_handles_2d() -> None:
    dt = parse_since("2d")
    assert (datetime.now(UTC) - dt) >= timedelta(days=2)
    assert (datetime.now(UTC) - dt) <= timedelta(days=2, minutes=1)


@freeze_time("2026-05-22T10:00:00Z")
def test_parse_since_handles_yesterday() -> None:
    dt = parse_since("yesterday")
    assert dt.day == 21
