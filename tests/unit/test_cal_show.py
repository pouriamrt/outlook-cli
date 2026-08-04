"""Tests for outlook cal show — get_event + ICS export."""

from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import MagicMock

from outlook_cli.graph.calendar import get_event

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_get_event_returns_parsed_event() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = json.loads(
        (FIXTURES / "graph_event.json").read_text()
    )
    ev = get_event(client, event_id="EV-1")
    assert ev.subject == "Team sync"
    assert ev.is_online_meeting is True


def test_get_event_uses_event_endpoint() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = json.loads(
        (FIXTURES / "graph_event.json").read_text()
    )
    get_event(client, event_id="EV-1")
    args, _ = client.get.call_args
    assert "/me/events/EV-1" in args[0]
