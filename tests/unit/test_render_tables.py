from datetime import UTC, datetime

from rich.console import Console

from outlook_cli.graph.models import Message, Recipient
from outlook_cli.render.tables import render_mail_list


def _msg(idx: int, subj: str, unread: bool = True) -> Message:
    return Message(
        id=f"ID-{idx}",
        subject=subj,
        preview=f"preview {idx}",
        importance="normal",
        is_read=not unread,
        is_flagged=False,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        conversation_id=f"CV-{idx}",
        **{"from": Recipient(name=f"Sender{idx}", address=f"s{idx}@example.com")},
        to=[Recipient(name="Pouria", address="pouriamortezaagha7@gmail.com")],
    )


def test_renders_mail_list_with_index_column(snapshot) -> None:
    messages = [_msg(1, "First"), _msg(2, "Second", unread=False)]
    out = Console(record=True, width=120)
    render_mail_list(out, messages)
    assert out.export_text(clear=True) == snapshot(name="mail_list_basic")


def test_renders_empty_list_gracefully() -> None:
    out = Console(record=True, width=120)
    render_mail_list(out, [])
    text = out.export_text()
    assert "No messages" in text or "0" in text
