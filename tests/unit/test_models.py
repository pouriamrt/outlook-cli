import json
from pathlib import Path

from outlook_cli.graph.models import Event, Folder, Message, Recipient

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def _load(name: str) -> dict:
    return json.loads((FIXTURES / name).read_text())


def test_recipient_parses_graph_shape() -> None:
    r = Recipient.from_graph({"emailAddress": {"name": "Alice", "address": "alice@example.com"}})
    assert r.name == "Alice"
    assert r.address == "alice@example.com"


def test_recipient_handles_missing_name() -> None:
    r = Recipient.from_graph({"emailAddress": {"address": "bob@example.com"}})
    assert r.name == ""
    assert r.address == "bob@example.com"


def test_message_round_trip_from_fixture() -> None:
    msg = Message.from_graph(_load("graph_message.json"))
    assert msg.id.startswith("AAMkAGI2")
    assert msg.subject == "Re: Q3 planning"
    assert msg.from_.address == "alice@example.com"
    assert msg.to[0].address == "pouriamortezaagha7@gmail.com"
    assert msg.is_read is False
    assert msg.is_flagged is True
    assert msg.has_attachments is True
    assert msg.importance == "normal"


def test_message_to_json_shape_matches_contract() -> None:
    msg = Message.from_graph(_load("graph_message.json"))
    out = msg.to_json_shape(index=1)
    expected_keys = {
        "id",
        "index",
        "from",
        "to",
        "subject",
        "received_at",
        "is_read",
        "is_flagged",
        "has_attachments",
        "importance",
        "preview",
        "conversation_id",
    }
    assert expected_keys.issubset(out.keys())
    assert out["index"] == 1
    assert out["from"] == {"name": "Alice", "address": "alice@example.com"}


def test_event_round_trip_from_fixture() -> None:
    ev = Event.from_graph(_load("graph_event.json"))
    assert ev.subject == "Team sync"
    assert ev.is_online_meeting is True
    assert ev.location == "Online"
    assert ev.attendees[0].address == "alice@example.com"


def test_event_to_json_shape() -> None:
    ev = Event.from_graph(_load("graph_event.json"))
    out = ev.to_json_shape(index=2)
    assert out["index"] == 2
    assert "start" in out and "end" in out
    assert out["online_meeting_url"].startswith("https://teams.microsoft.com")


def test_folder_round_trip() -> None:
    f = Folder.from_graph(
        {"id": "AAMkF==", "displayName": "Inbox", "totalItemCount": 142, "unreadItemCount": 7}
    )
    assert f.display_name == "Inbox"
    assert f.unread == 7
    assert f.total == 142
