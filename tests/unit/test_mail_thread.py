"""Unit tests for read_thread."""

from __future__ import annotations

from unittest.mock import MagicMock

from outlook_cli.graph.mail import read_thread


def test_read_thread_filters_by_conversation_id_and_orders() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {
                "id": "M2",
                "subject": "Re",
                "isRead": True,
                "receivedDateTime": "2026-05-22T11:00:00Z",
                "conversationId": "CV",
                "from": {"emailAddress": {"name": "B", "address": "b@c"}},
                "body": {"contentType": "text", "content": "second"},
            },
            {
                "id": "M1",
                "subject": "Re",
                "isRead": True,
                "receivedDateTime": "2026-05-22T10:00:00Z",
                "conversationId": "CV",
                "from": {"emailAddress": {"name": "A", "address": "a@c"}},
                "body": {"contentType": "text", "content": "first"},
            },
        ]
    }
    msgs = read_thread(client, conversation_id="CV")
    args, kwargs = client.get.call_args
    assert "conversationId eq 'CV'" in kwargs["params"]["$filter"]
    assert msgs[0].body_html == "first"
    assert msgs[1].body_html == "second"
