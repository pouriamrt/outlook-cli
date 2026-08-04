"""ICS escape + line-folding tests for graph/calendar.event_to_ics security fix."""

from datetime import UTC, datetime

from outlook_cli.graph.calendar import _ics_escape, _ics_fold, event_to_ics
from outlook_cli.graph.models import Attendee, Event, Recipient


def test_ics_escape_handles_crlf_in_text() -> None:
    assert _ics_escape("a\r\nb") == "a\\nb"
    assert _ics_escape("a\nb") == "a\\nb"
    assert _ics_escape("a\rb") == "a\\nb"


def test_ics_escape_handles_metacharacters() -> None:
    assert _ics_escape("a,b") == "a\\,b"
    assert _ics_escape("a;b") == "a\\;b"
    assert _ics_escape("a\\b") == "a\\\\b"


def test_ics_escape_backslash_order_is_correct() -> None:
    # Backslash must be escaped FIRST so subsequent replacements don't double-escape.
    # Input: literal backslash followed by newline
    assert _ics_escape("\\\n") == "\\\\\\n"


def test_ics_fold_short_line_untouched() -> None:
    short = "BEGIN:VEVENT"
    assert _ics_fold(short) == short


def test_ics_fold_long_ascii_line_wraps_at_75_octets() -> None:
    line = "SUMMARY:" + "x" * 200
    folded = _ics_fold(line)
    assert "\r\n " in folded
    for chunk in folded.split("\r\n "):
        assert len(chunk.encode("utf-8")) <= 75


def test_ics_fold_preserves_utf8_multibyte_boundaries() -> None:
    # Each emoji is 4 UTF-8 bytes; place enough to force a fold across one
    line = "SUMMARY:" + ("\U0001f680" * 30)  # 30 rockets = 120 bytes
    folded = _ics_fold(line)
    # Folded back together should equal the original
    reassembled = folded.replace("\r\n ", "")
    assert reassembled == line


def test_event_to_ics_neutralizes_crlf_injection_in_subject() -> None:
    ev = Event(
        id="EV-1",
        subject="Normal subject\r\nDTEND:19700101T000000Z\r\nSUMMARY:Injected",
        start=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, 10, 30, 0, tzinfo=UTC),
        organizer=Recipient(name="P", address="p@b"),
        attendees=[],
    )
    out = event_to_ics(ev)
    # The injected SUMMARY line must NOT appear as a real property
    assert "\r\nSUMMARY:Injected" not in out
    # The CRLF in subject must have been collapsed to \n escape
    assert "\\n" in out


def test_event_to_ics_neutralizes_injection_in_organizer_name() -> None:
    ev = Event(
        id="EV-1",
        subject="Sync",
        start=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, 10, 30, 0, tzinfo=UTC),
        organizer=Recipient(
            name="Alice\r\nATTENDEE;CN=mallory:mailto:m@evil",
            address="alice@example.com",
        ),
        attendees=[],
    )
    out = event_to_ics(ev)
    # Injected ATTENDEE must not appear as a real property line
    assert "\r\nATTENDEE;CN=mallory" not in out


def test_event_to_ics_escapes_attendee_fields() -> None:
    ev = Event(
        id="EV-1",
        subject="Sync",
        start=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        end=datetime(2026, 5, 22, 10, 30, 0, tzinfo=UTC),
        organizer=Recipient(name="P", address="p@b"),
        attendees=[
            Attendee(
                name="Bob\r\nDTEND:19700101T000000Z",
                address="bob@example.com",
                required=True,
                response="accepted",
            )
        ],
    )
    out = event_to_ics(ev)
    assert "\r\nDTEND:1970" not in out
