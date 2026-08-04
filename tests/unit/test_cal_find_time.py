from datetime import UTC, date, datetime, time, timedelta
from unittest.mock import MagicMock

from outlook_cli.commands.calendar import _business_days, _day_window, _spread
from outlook_cli.graph.calendar import find_meeting_times


def test_find_meeting_times_posts_with_attendees_and_window() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"meetingTimeSuggestions": []}
    start = datetime(2026, 5, 22, 9, 0, 0, tzinfo=UTC)
    end = datetime(2026, 5, 22, 17, 0, 0, tzinfo=UTC)
    find_meeting_times(
        client,
        attendees=["alice@example.com", "bob@example.com"],
        duration=timedelta(minutes=30),
        window_start=start,
        window_end=end,
        max_candidates=5,
    )
    args, kwargs = client.post.call_args
    assert "/me/findMeetingTimes" in args[0]
    body = kwargs["json_body"]
    assert len(body["attendees"]) == 2
    assert body["meetingDuration"] == "PT30M"
    assert body["maxCandidates"] == 5
    slot_start = body["timeConstraint"]["timeslots"][0]["start"]["dateTime"]
    assert slot_start.startswith("2026-05-22T09:00")


def test_find_meeting_times_returns_suggestions() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {
        "meetingTimeSuggestions": [
            {
                "meetingTimeSlot": {
                    "start": {"dateTime": "2026-05-22T10:00:00.0000000", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-05-22T10:30:00.0000000", "timeZone": "UTC"},
                },
                "confidence": 90.0,
                "organizerAvailability": "free",
                "attendeeAvailability": [],
                "locations": [],
            }
        ]
    }
    suggestions = find_meeting_times(
        client,
        attendees=["a@b"],
        duration=timedelta(minutes=30),
        window_start=datetime(2026, 5, 22, tzinfo=UTC),
        window_end=datetime(2026, 5, 22, 23, 59, tzinfo=UTC),
        max_candidates=3,
    )
    assert len(suggestions) == 1
    assert suggestions[0].confidence == 90.0


# ---------- window resolution: _business_days ----------

_TUE = date(2026, 5, 26)  # a Tuesday
_SAT = date(2026, 5, 30)  # a Saturday


def test_business_days_next_week_is_next_monday_through_friday() -> None:
    days = _business_days("next week", today=_TUE)
    assert days == [
        date(2026, 6, 1),
        date(2026, 6, 2),
        date(2026, 6, 3),
        date(2026, 6, 4),
        date(2026, 6, 5),
    ]
    assert days[0].weekday() == 0  # Monday, not today+7 (which was the bug)


def test_business_days_this_week_skips_past_days() -> None:
    days = _business_days("this week", today=_TUE)
    # This calendar week's Mon-Fri = 25..29; only days >= today (Tue 26) kept.
    assert days == [date(2026, 5, 26), date(2026, 5, 27), date(2026, 5, 28), date(2026, 5, 29)]


def test_business_days_this_week_on_weekend_falls_back_to_next_week() -> None:
    days = _business_days("this week", today=_SAT)
    assert days[0] == date(2026, 6, 1)  # next Monday
    assert len(days) == 5


def test_business_days_specific_date_is_single_day() -> None:
    days = _business_days("2026-06-10", today=_TUE)
    assert days == [date(2026, 6, 10)]


def test_business_days_next_month_does_not_match_next_week() -> None:
    # "next month" contains the substring "next" but is NOT the "next week"
    # keyword. Substring matching wrongly returned a 5-day week; exact matching
    # routes it to single-day parsing.
    days = _business_days("next month", today=_TUE)
    assert len(days) == 1


# ---------- per-day window: _day_window ----------


def test_day_window_future_day_uses_working_hours() -> None:
    now = datetime.combine(_TUE, time(14, 0)).astimezone()
    win = _day_window(date(2026, 6, 1), now=now)
    assert win is not None
    start, end = win
    assert (start.hour, start.minute) == (9, 0)
    assert (end.hour, end.minute) == (17, 0)


def test_day_window_today_clamps_start_to_now() -> None:
    now = datetime.combine(_TUE, time(14, 0)).astimezone()
    win = _day_window(_TUE, now=now)
    assert win is not None
    start, end = win
    assert start == now  # don't suggest past slots earlier today
    assert (end.hour, end.minute) == (17, 0)


def test_day_window_fully_past_day_returns_none() -> None:
    now = datetime.combine(_TUE, time(14, 0)).astimezone()
    assert _day_window(date(2026, 5, 20), now=now) is None


# ---------- timezone parsing: Graph returns offset-naive UTC + timeZone field ----------


def test_find_meeting_times_treats_naive_datetime_as_utc() -> None:
    """Graph emits 'dateTime' with no offset + a sibling timeZone:'UTC'.

    The naive string must be parsed as UTC, not local, or every slot renders
    hours-shifted (e.g. 09:00 UTC morning slots displayed as afternoon).
    """
    client = MagicMock()
    client.post.return_value.json.return_value = {
        "meetingTimeSuggestions": [
            {
                "meetingTimeSlot": {
                    "start": {"dateTime": "2026-05-27T13:00:00.0000000", "timeZone": "UTC"},
                    "end": {"dateTime": "2026-05-27T13:30:00.0000000", "timeZone": "UTC"},
                },
                "confidence": 100.0,
            }
        ]
    }
    suggestions = find_meeting_times(
        client,
        attendees=["a@b"],
        duration=timedelta(minutes=30),
        window_start=datetime(2026, 5, 27, tzinfo=UTC),
        window_end=datetime(2026, 5, 27, 23, 59, tzinfo=UTC),
        max_candidates=3,
    )
    s = suggestions[0]
    assert s.start.tzinfo is not None
    assert s.start.utcoffset() == timedelta(0)  # UTC-aware
    assert s.start.astimezone(UTC).hour == 13  # 13:00 UTC, not shifted to local


def test_event_from_graph_treats_naive_datetime_as_utc() -> None:
    from outlook_cli.graph.models import Event

    ev = Event.from_graph(
        {
            "id": "E1",
            "subject": "x",
            "start": {"dateTime": "2026-05-27T13:00:00.0000000", "timeZone": "UTC"},
            "end": {"dateTime": "2026-05-27T13:30:00.0000000", "timeZone": "UTC"},
        }
    )
    assert ev.start.utcoffset() == timedelta(0)
    assert ev.start.astimezone(UTC).hour == 13


# ---------- within-day spread: _spread ----------


def test_spread_returns_all_when_fewer_than_k() -> None:
    assert _spread([1, 2], 3) == [1, 2]


def test_spread_includes_first_and_last() -> None:
    out = _spread(list(range(16)), 3)
    assert len(out) == 3
    assert out[0] == 0  # earliest slot
    assert out[-1] == 15  # latest slot (afternoon), not just the earliest run


def test_spread_samples_evenly() -> None:
    # 8 slots, want 3 spread → first, middle, last (not the earliest 3).
    assert _spread(list(range(8)), 3) == [0, 4, 7]


def test_spread_single() -> None:
    assert _spread(list(range(10)), 1) == [0]
