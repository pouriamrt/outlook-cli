from datetime import UTC, datetime

from rich.console import Console

from outlook_cli.graph.models import Message, Recipient
from outlook_cli.render.detail import render_message_detail


def _msg() -> Message:
    return Message(
        id="MSG",
        subject="Hello",
        preview="hi",
        importance="high",
        is_read=False,
        is_flagged=True,
        has_attachments=False,
        received_at=datetime(2026, 5, 22, 10, 0, 0, tzinfo=UTC),
        conversation_id="CV",
        **{"from": Recipient(name="Alice", address="alice@example.com")},
        to=[Recipient(name="Pouria", address="pouriamortezaagha7@gmail.com")],
        body_html="<p>Hello <b>world</b></p>",
        body_content_type="html",
    )


def test_renders_headers_and_body() -> None:
    out = Console(record=True, width=120)
    render_message_detail(out, _msg(), raw=False)
    text = out.export_text()
    assert "Alice" in text
    assert "alice@example.com" in text
    assert "Hello" in text
    assert "world" in text


def test_raw_mode_does_not_convert_html() -> None:
    out = Console(record=True, width=120)
    render_message_detail(out, _msg(), raw=True)
    text = out.export_text()
    assert "<p>" in text or "<b>" in text


def test_text_content_type_passes_through() -> None:
    msg = _msg()
    msg.body_html = "Plain text body."
    msg.body_content_type = "text"
    out = Console(record=True, width=120)
    render_message_detail(out, msg, raw=False)
    text = out.export_text()
    assert "Plain text body." in text
