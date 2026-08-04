"""Pydantic models for Microsoft Graph entities. Provides Graph→model and model→JSON shape."""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class Recipient(BaseModel):
    model_config = ConfigDict(populate_by_name=True)
    name: str = ""
    address: str = ""

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> Recipient:
        email = payload.get("emailAddress") or {}
        return cls(name=email.get("name", "") or "", address=email.get("address", "") or "")

    def to_json_shape(self) -> dict[str, str]:
        return {"name": self.name, "address": self.address}


def _parse_iso(s: str) -> datetime:
    if not s:
        raise ValueError("ISO datetime string is empty")
    s = s.replace("Z", "+00:00")
    return datetime.fromisoformat(s)


def _parse_graph_dt(node: dict[str, Any]) -> datetime:
    """Parse a Graph ``{dateTime, timeZone}`` node to a tz-aware datetime.

    Graph (calendarView, findMeetingTimes) emits offset-naive strings such as
    ``'2026-05-27T13:00:00.0000000'`` with the zone carried in a sibling
    ``timeZone`` field that defaults to ``UTC``. Parsing the string alone yields
    a naive datetime, and a later ``.astimezone()`` would then mislabel that UTC
    instant as local time — shifting every value by the local offset. We attach
    UTC to naive values so downstream conversion is correct.
    """
    raw = (node or {}).get("dateTime", "")
    if not raw:
        raise ValueError("ISO datetime string is empty")
    dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    return dt


class Message(BaseModel):
    model_config = ConfigDict(populate_by_name=True)

    id: str
    subject: str = ""
    preview: str = ""
    importance: str = "normal"
    is_read: bool = False
    is_flagged: bool = False
    has_attachments: bool = False
    received_at: datetime
    conversation_id: str = ""
    from_: Recipient = Field(default_factory=Recipient, alias="from")
    to: list[Recipient] = Field(default_factory=list)
    cc: list[Recipient] = Field(default_factory=list)
    body_html: str = ""
    body_content_type: str = "text"

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> Message:
        flag = (payload.get("flag") or {}).get("flagStatus", "notFlagged")
        body = payload.get("body") or {}
        received_str = payload.get("receivedDateTime", "") or ""
        received_at = _parse_iso(received_str) if received_str else datetime.now(UTC)
        return cls(
            id=payload["id"],
            subject=payload.get("subject", "") or "",
            preview=payload.get("bodyPreview", "") or "",
            importance=payload.get("importance", "normal"),
            is_read=bool(payload.get("isRead", False)),
            is_flagged=flag == "flagged",
            has_attachments=bool(payload.get("hasAttachments", False)),
            received_at=received_at,
            conversation_id=payload.get("conversationId", "") or "",
            **{"from": Recipient.from_graph(payload.get("from") or {})},
            to=[Recipient.from_graph(r) for r in (payload.get("toRecipients") or [])],
            cc=[Recipient.from_graph(r) for r in (payload.get("ccRecipients") or [])],
            body_html=body.get("content", "") or "",
            body_content_type=body.get("contentType", "text"),
        )

    def to_json_shape(self, *, index: int) -> dict[str, Any]:
        return {
            "id": self.id,
            "index": index,
            "from": self.from_.to_json_shape(),
            "to": [r.to_json_shape() for r in self.to],
            "cc": [r.to_json_shape() for r in self.cc],
            "subject": self.subject,
            "received_at": self.received_at.isoformat(),
            "is_read": self.is_read,
            "is_flagged": self.is_flagged,
            "has_attachments": self.has_attachments,
            "importance": self.importance,
            "preview": self.preview,
            "conversation_id": self.conversation_id,
        }


class Attendee(BaseModel):
    name: str = ""
    address: str = ""
    response: str = "none"
    required: bool = True

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> Attendee:
        email = payload.get("emailAddress") or {}
        status = payload.get("status") or {}
        return cls(
            name=email.get("name", "") or "",
            address=email.get("address", "") or "",
            response=status.get("response", "none"),
            required=payload.get("type", "required") == "required",
        )

    def to_json_shape(self) -> dict[str, Any]:
        return {
            "name": self.name,
            "address": self.address,
            "response": self.response,
            "required": self.required,
        }


class Event(BaseModel):
    id: str
    subject: str = ""
    preview: str = ""
    start: datetime
    end: datetime
    is_all_day: bool = False
    is_online_meeting: bool = False
    online_meeting_url: str = ""
    location: str = ""
    organizer: Recipient = Field(default_factory=Recipient)
    attendees: list[Attendee] = Field(default_factory=list)
    response_status: str = "none"
    body_html: str = ""

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> Event:
        body = payload.get("body") or {}
        location = (payload.get("location") or {}).get("displayName", "") or ""
        return cls(
            id=payload["id"],
            subject=payload.get("subject", "") or "",
            preview=payload.get("bodyPreview", "") or "",
            start=_parse_graph_dt(payload.get("start") or {}),
            end=_parse_graph_dt(payload.get("end") or {}),
            is_all_day=bool(payload.get("isAllDay", False)),
            is_online_meeting=bool(payload.get("isOnlineMeeting", False)),
            online_meeting_url=payload.get("onlineMeetingUrl", "") or "",
            location=location,
            organizer=Recipient.from_graph(payload.get("organizer") or {}),
            attendees=[Attendee.from_graph(a) for a in (payload.get("attendees") or [])],
            response_status=(payload.get("responseStatus") or {}).get("response", "none"),
            body_html=body.get("content", "") or "",
        )

    def to_json_shape(self, *, index: int) -> dict[str, Any]:
        return {
            "id": self.id,
            "index": index,
            "subject": self.subject,
            "start": self.start.isoformat(),
            "end": self.end.isoformat(),
            "is_all_day": self.is_all_day,
            "is_online_meeting": self.is_online_meeting,
            "online_meeting_url": self.online_meeting_url,
            "location": self.location,
            "organizer": self.organizer.to_json_shape(),
            "attendees": [a.to_json_shape() for a in self.attendees],
            "response_status": self.response_status,
            "preview": self.preview,
        }


class Folder(BaseModel):
    id: str
    display_name: str
    total: int = 0
    unread: int = 0

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> Folder:
        return cls(
            id=payload["id"],
            display_name=payload.get("displayName", "") or "",
            total=int(payload.get("totalItemCount", 0)),
            unread=int(payload.get("unreadItemCount", 0)),
        )


class FindTimeSuggestion(BaseModel):
    start: datetime
    end: datetime
    confidence: float = 0.0
    organizer_availability: str = "unknown"
    attendee_availability: list[dict[str, str]] = Field(default_factory=list)
    locations: list[str] = Field(default_factory=list)

    @classmethod
    def from_graph(cls, payload: dict[str, Any]) -> FindTimeSuggestion:
        slot = payload.get("meetingTimeSlot") or {}
        return cls(
            start=_parse_graph_dt(slot.get("start") or {}),
            end=_parse_graph_dt(slot.get("end") or {}),
            confidence=float(payload.get("confidence", 0.0)),
            organizer_availability=payload.get("organizerAvailability", "unknown"),
            attendee_availability=[
                {
                    "address": (a.get("attendee") or {}).get("emailAddress", {}).get("address", ""),
                    "availability": a.get("availability", "unknown"),
                }
                for a in (payload.get("attendeeAvailability") or [])
            ],
            locations=[(loc.get("displayName") or "") for loc in (payload.get("locations") or [])],
        )

    def to_json_shape(self) -> dict[str, Any]:
        return self.model_dump(mode="json")
