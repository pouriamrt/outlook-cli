from datetime import UTC, datetime
from unittest.mock import MagicMock

from outlook_cli.graph.calendar import list_events


def test_list_events_calls_calendarView_with_range() -> None:  # noqa: N802
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": []}
    start = datetime(2026, 5, 22, 0, 0, 0, tzinfo=UTC)
    end = datetime(2026, 5, 23, 0, 0, 0, tzinfo=UTC)
    list_events(client, start=start, end=end, calendar_name=None)
    args, kwargs = client.get.call_args
    assert "/me/calendarView" in args[0]
    assert kwargs["params"]["startDateTime"] == "2026-05-22T00:00:00Z"
    assert kwargs["params"]["endDateTime"] == "2026-05-23T00:00:00Z"


def test_list_events_returns_parsed_events() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {
                "id": "EV1",
                "subject": "Sync",
                "start": {"dateTime": "2026-05-22T10:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-05-22T10:30:00", "timeZone": "UTC"},
                "organizer": {"emailAddress": {"name": "P", "address": "p@b"}},
                "attendees": [],
                "responseStatus": {"response": "organizer", "time": "0001-01-01T00:00:00Z"},
            }
        ]
    }
    events = list_events(
        client, start=datetime(2026, 5, 22), end=datetime(2026, 5, 23), calendar_name=None
    )
    assert len(events) == 1
    assert events[0].subject == "Sync"


def test_list_events_uses_named_calendar_endpoint() -> None:
    client = MagicMock()
    client.get.side_effect = [
        MagicMock(json=lambda: {"value": [{"id": "CAL-PERSONAL", "name": "Personal"}]}),
        MagicMock(json=lambda: {"value": []}),
    ]
    list_events(
        client, start=datetime(2026, 5, 22), end=datetime(2026, 5, 23), calendar_name="Personal"
    )
    second_path = client.get.call_args_list[1].args[0]
    assert "/me/calendars/CAL-PERSONAL/calendarView" in second_path
