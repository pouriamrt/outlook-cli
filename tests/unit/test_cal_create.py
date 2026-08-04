from datetime import UTC, datetime
from unittest.mock import MagicMock

from outlook_cli.graph.calendar import CreateEvent, create_event


def _create() -> CreateEvent:
    return CreateEvent(
        title="Sync",
        start=datetime(2026, 5, 23, 15, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 23, 15, 30, 0, tzinfo=UTC),
        invitees=["alice@example.com"],
        location="Room 1",
        body="<p>Agenda</p>",
        is_online_meeting=False,
        is_all_day=False,
    )


def test_create_event_posts_to_events_endpoint() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "EV-NEW"}
    eid = create_event(client, _create())
    args, kwargs = client.post.call_args
    assert "/me/events" in args[0]
    assert eid == "EV-NEW"


def test_create_event_sets_subject_and_times() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "EV"}
    create_event(client, _create())
    body = client.post.call_args.kwargs["json_body"]
    assert body["subject"] == "Sync"
    assert body["start"]["dateTime"] == "2026-05-23T15:00:00"
    assert body["end"]["dateTime"] == "2026-05-23T15:30:00"
    assert body["start"]["timeZone"] == "UTC"


def test_create_event_includes_attendees() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "EV"}
    create_event(client, _create())
    body = client.post.call_args.kwargs["json_body"]
    assert body["attendees"][0]["emailAddress"]["address"] == "alice@example.com"


def test_create_event_online_meeting_true() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "EV"}
    spec = _create()
    spec.is_online_meeting = True
    create_event(client, spec)
    body = client.post.call_args.kwargs["json_body"]
    assert body["isOnlineMeeting"] is True
    assert body["onlineMeetingProvider"] == "teamsForBusiness"


def test_create_event_all_day_uses_date_format() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "EV"}
    spec = _create()
    spec.is_all_day = True
    spec.start = datetime(2026, 5, 23, tzinfo=UTC)
    spec.end = datetime(2026, 5, 24, tzinfo=UTC)
    create_event(client, spec)
    body = client.post.call_args.kwargs["json_body"]
    assert body["isAllDay"] is True
