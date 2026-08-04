import json
from datetime import UTC, datetime

from outlook_cli.graph.models import Message, Recipient
from outlook_cli.render.json_out import dump_mail_list


def _msg(idx: int) -> Message:
    return Message(
        id=f"ID-{idx}",
        subject=f"Subject {idx}",
        preview="preview",
        importance="normal",
        is_read=False,
        is_flagged=False,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        conversation_id=f"CV-{idx}",
        **{"from": Recipient(name="A", address="a@b.ca")},
        to=[Recipient(name="B", address="b@c.ca")],
    )


def test_dump_mail_list_emits_stable_keys() -> None:
    out = dump_mail_list([_msg(1), _msg(2)], next_skip=25, total_estimated=None)
    payload = json.loads(out)
    assert set(payload.keys()) >= {"items", "next_skip"}
    item0 = payload["items"][0]
    assert set(item0.keys()) >= {
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


def test_dump_mail_list_indices_are_1_based() -> None:
    payload = json.loads(dump_mail_list([_msg(1), _msg(2)], next_skip=None))
    assert payload["items"][0]["index"] == 1
    assert payload["items"][1]["index"] == 2
