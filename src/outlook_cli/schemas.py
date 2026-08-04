"""JSON Schemas for --json command outputs. Frozen after v1.0."""

from __future__ import annotations

from typing import Any

_RECIPIENT: dict[str, Any] = {
    "type": "object",
    "properties": {"name": {"type": "string"}, "address": {"type": "string"}},
    "required": ["name", "address"],
}

_MAIL_ITEM: dict[str, Any] = {
    "type": "object",
    "properties": {
        "id": {"type": "string"},
        "index": {"type": "integer"},
        "from": _RECIPIENT,
        "to": {"type": "array", "items": _RECIPIENT},
        "subject": {"type": "string"},
        "received_at": {"type": "string", "format": "date-time"},
        "is_read": {"type": "boolean"},
        "is_flagged": {"type": "boolean"},
        "has_attachments": {"type": "boolean"},
        "importance": {"type": "string"},
        "preview": {"type": "string"},
        "conversation_id": {"type": "string"},
    },
    "required": ["id", "subject", "received_at", "is_read"],
}

SCHEMAS: dict[str, dict[str, Any]] = {
    "mail.list": {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "title": "outlook mail list",
        "type": "object",
        "properties": {
            "items": {"type": "array", "items": _MAIL_ITEM},
            "next_skip": {"type": ["integer", "null"]},
            "total_estimated": {"type": ["integer", "null"]},
        },
        "required": ["items"],
    },
    "mail.read": {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "title": "outlook mail read",
        "allOf": [
            _MAIL_ITEM,
            {
                "type": "object",
                "properties": {
                    "body_html": {"type": "string"},
                    "body_content_type": {
                        "type": "string",
                        "enum": ["html", "text"],
                    },
                },
            },
        ],
    },
    "cal.list": {
        "$schema": "https://json-schema.org/draft/2020-12/schema",
        "title": "outlook cal list",
        "type": "object",
        "properties": {
            "items": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "id": {"type": "string"},
                        "index": {"type": "integer"},
                        "subject": {"type": "string"},
                        "start": {"type": "string", "format": "date-time"},
                        "end": {"type": "string", "format": "date-time"},
                        "is_all_day": {"type": "boolean"},
                        "is_online_meeting": {"type": "boolean"},
                        "location": {"type": "string"},
                    },
                    "required": ["id", "subject", "start", "end"],
                },
            },
        },
    },
}


def get_schema(name: str) -> dict[str, Any] | None:
    return SCHEMAS.get(name)
