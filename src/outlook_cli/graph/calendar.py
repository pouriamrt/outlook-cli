"""Calendar primitives on Microsoft Graph."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, datetime, timedelta
from typing import Any

from outlook_cli.errors import NotFound
from outlook_cli.graph.client import GraphClient
from outlook_cli.graph.models import Event, FindTimeSuggestion

_EVENT_SELECT = (
    "id,subject,bodyPreview,body,start,end,isAllDay,isOnlineMeeting,"
    "onlineMeetingUrl,location,organizer,attendees,responseStatus"
)


def _iso(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def _resolve_calendar_id(client: GraphClient, name: str) -> str:
    resp = client.get("/me/calendars", params={"$top": 100, "$select": "id,name"})
    for cal in resp.json().get("value", []):
        if cal["name"].lower() == name.lower():
            cal_id: str = cal["id"]
            return cal_id
    raise NotFound(f"Calendar '{name}' not found.")


def list_events(
    client: GraphClient,
    *,
    start: datetime,
    end: datetime,
    calendar_name: str | None,
) -> list[Event]:
    if calendar_name:
        cal_id = _resolve_calendar_id(client, calendar_name)
        path = f"/me/calendars/{cal_id}/calendarView"
    else:
        path = "/me/calendarView"
    resp = client.get(
        path,
        params={
            "startDateTime": _iso(start),
            "endDateTime": _iso(end),
            "$top": 250,
            "$orderby": "start/dateTime",
            "$select": _EVENT_SELECT,
        },
    )
    return [Event.from_graph(e) for e in resp.json().get("value", [])]


def get_event(client: GraphClient, *, event_id: str) -> Event:
    resp = client.get(
        f"/me/events/{event_id}",
        params={"$select": _EVENT_SELECT},
    )
    return Event.from_graph(resp.json())


def _ics_escape(value: str) -> str:
    """RFC 5545 §3.3.11 TEXT escaping: CRLF, semicolons, commas, backslashes."""
    return (
        value.replace("\\", "\\\\")
        .replace("\r\n", "\\n")
        .replace("\r", "\\n")
        .replace("\n", "\\n")
        .replace(",", "\\,")
        .replace(";", "\\;")
    )


def _ics_fold(line: str) -> str:
    """RFC 5545 §3.1 line folding at 75 octets (UTF-8)."""
    encoded = line.encode("utf-8")
    if len(encoded) <= 75:
        return line
    chunks: list[bytes] = []
    while len(encoded) > 75:
        # Find a safe split point that doesn't break a multibyte char
        split = 75
        # Multi-byte UTF-8 chars are at most 4 bytes, so we backtrack at most 3
        # positions to find a code-point boundary. The loop terminator is safe.
        while split > 0 and (encoded[split] & 0xC0) == 0x80:
            split -= 1
        chunks.append(encoded[:split])
        encoded = encoded[split:]
    chunks.append(encoded)
    return "\r\n ".join(c.decode("utf-8") for c in chunks)


def event_to_ics(ev: Event) -> str:
    """Minimal RFC 5545 .ics serialization (UTC times, organizer + attendees + summary).

    Applies TEXT escaping to user-supplied fields and folds lines at 75 octets.
    """

    def _fmt(dt: datetime) -> str:
        return dt.astimezone(UTC).strftime("%Y%m%dT%H%M%SZ")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//outlook-cli//EN",
        "BEGIN:VEVENT",
        f"UID:{_ics_escape(ev.id)}",
        f"DTSTAMP:{_fmt(ev.start)}",
        f"DTSTART:{_fmt(ev.start)}",
        f"DTEND:{_fmt(ev.end)}",
        f"SUMMARY:{_ics_escape(ev.subject or '')}",
        f"ORGANIZER;CN={_ics_escape(ev.organizer.name)}:mailto:{_ics_escape(ev.organizer.address)}",
    ]
    for a in ev.attendees:
        role = "REQ" if a.required else "OPT"
        lines.append(
            f"ATTENDEE;CN={_ics_escape(a.name)};ROLE={role}-PARTICIPANT:mailto:{_ics_escape(a.address)}"
        )
    if ev.location:
        lines.append(f"LOCATION:{_ics_escape(ev.location)}")
    lines.extend(["END:VEVENT", "END:VCALENDAR"])
    return "\r\n".join(_ics_fold(line) for line in lines)


@dataclass
class CreateEvent:
    title: str
    start: datetime
    end: datetime
    invitees: list[str] = field(default_factory=list)
    location: str = ""
    body: str = ""
    is_online_meeting: bool = False
    is_all_day: bool = False


def create_event(client: GraphClient, spec: CreateEvent) -> str:
    payload: dict[str, Any] = {
        "subject": spec.title,
        "start": {
            "dateTime": spec.start.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": spec.end.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": "UTC",
        },
        "isAllDay": spec.is_all_day,
        "attendees": [{"emailAddress": {"address": a}, "type": "required"} for a in spec.invitees],
    }
    if spec.location:
        payload["location"] = {"displayName": spec.location}
    if spec.body:
        payload["body"] = {"contentType": "HTML", "content": spec.body}
    if spec.is_online_meeting:
        payload["isOnlineMeeting"] = True
        payload["onlineMeetingProvider"] = "teamsForBusiness"
    resp = client.post("/me/events", json_body=payload)
    event_id: str = resp.json()["id"]
    return event_id


_RESPOND_MAP = {"accept": "accept", "decline": "decline", "tentative": "tentativelyAccept"}


def respond_event(client: GraphClient, *, event_id: str, verb: str, comment: str) -> None:
    endpoint = _RESPOND_MAP.get(verb)
    if endpoint is None:
        raise ValueError(f"Unknown response verb: {verb!r}")
    client.post(
        f"/me/events/{event_id}/{endpoint}",
        json_body={"comment": comment, "sendResponse": True},
    )


def cancel_event(client: GraphClient, *, event_id: str, comment: str) -> None:
    client.post(f"/me/events/{event_id}/cancel", json_body={"Comment": comment})


def _iso_local(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S")


def _duration_to_iso(td: timedelta) -> str:
    total = int(td.total_seconds())
    hours, rem = divmod(total, 3600)
    minutes = rem // 60
    parts: list[str] = []
    if hours:
        parts.append(f"{hours}H")
    if minutes:
        parts.append(f"{minutes}M")
    if not parts:
        parts.append("0M")
    return "PT" + "".join(parts)


def find_meeting_times(
    client: GraphClient,
    *,
    attendees: list[str],
    duration: timedelta,
    window_start: datetime,
    window_end: datetime,
    max_candidates: int = 5,
) -> list[FindTimeSuggestion]:
    body: dict[str, Any] = {
        "attendees": [{"emailAddress": {"address": a}, "type": "required"} for a in attendees],
        "timeConstraint": {
            "timeslots": [
                {
                    "start": {"dateTime": _iso_local(window_start), "timeZone": "UTC"},
                    "end": {"dateTime": _iso_local(window_end), "timeZone": "UTC"},
                }
            ],
        },
        "meetingDuration": _duration_to_iso(duration),
        "maxCandidates": max_candidates,
        "isOrganizerOptional": False,
        "returnSuggestionReasons": True,
        "minimumAttendeePercentage": 100,
    }
    resp = client.post("/me/findMeetingTimes", json_body=body)
    return [FindTimeSuggestion.from_graph(s) for s in resp.json().get("meetingTimeSuggestions", [])]
