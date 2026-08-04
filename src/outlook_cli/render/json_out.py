"""Stable JSON output for --json mode. Schema is versioned and must not break consumers."""

from __future__ import annotations

import json
from typing import Any

from outlook_cli.graph.models import Event, Message


def dump_mail_list(
    messages: list[Message],
    *,
    next_skip: int | None,
    total_estimated: int | None = None,
) -> str:
    items = [msg.to_json_shape(index=i) for i, msg in enumerate(messages, start=1)]
    payload: dict[str, Any] = {"items": items, "next_skip": next_skip}
    if total_estimated is not None:
        payload["total_estimated"] = total_estimated
    return json.dumps(payload, indent=2, default=str)


def dump_event_list(events: list[Event]) -> str:
    items = [ev.to_json_shape(index=i) for i, ev in enumerate(events, start=1)]
    return json.dumps({"items": items}, indent=2, default=str)


def dump_object(obj: dict[str, Any]) -> str:
    return json.dumps(obj, indent=2, default=str)


def dump_mail_detail(msg: Message) -> str:
    base = msg.to_json_shape(index=0)
    base["body_html"] = msg.body_html
    base["body_content_type"] = msg.body_content_type
    return json.dumps(base, indent=2, default=str)


def dump_event_detail(ev: Event) -> str:
    base = ev.to_json_shape(index=0)
    base["body_html"] = ev.body_html
    return json.dumps(base, indent=2, default=str)
