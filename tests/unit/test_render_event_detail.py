"""Unit tests for render_event_detail."""

from __future__ import annotations

from datetime import UTC, datetime

from rich.console import Console

from outlook_cli.graph.models import Attendee, Event, Recipient
from outlook_cli.render.detail import render_event_detail


def _ev() -> Event:
    return Event(
        id="EV-1",
        subject="Team sync",
        start=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, 10, 30, 0, tzinfo=UTC),
        organizer=Recipient(name="Alice", address="alice@example.com"),
        location="Room 1",
        is_online_meeting=True,
        online_meeting_url="https://teams.microsoft.com/x",
        attendees=[
            Attendee(name="Bob", address="bob@example.com", required=True, response="accepted")
        ],
        body_html="<p>Agenda</p>",
    )


def test_render_event_detail_basic() -> None:
    out = Console(record=True, width=120)
    render_event_detail(out, _ev(), show_attendees=False)
    text = out.export_text()
    assert "Team sync" in text
    assert "Room 1" in text
    assert "Alice" in text


def test_render_event_detail_with_attendees() -> None:
    out = Console(record=True, width=120)
    render_event_detail(out, _ev(), show_attendees=True)
    text = out.export_text()
    assert "Bob" in text
    assert "accepted" in text


def test_render_event_detail_includes_teams_link() -> None:
    out = Console(record=True, width=120)
    render_event_detail(out, _ev(), show_attendees=False)
    text = out.export_text()
    assert "teams.microsoft.com" in text


def test_render_event_detail_no_location_no_teams() -> None:
    ev = _ev()
    ev.location = ""
    ev.is_online_meeting = False
    ev.online_meeting_url = ""
    out = Console(record=True, width=120)
    render_event_detail(out, ev, show_attendees=False)
    text = out.export_text()
    assert "Room 1" not in text
    assert "teams.microsoft.com" not in text
