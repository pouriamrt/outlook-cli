import base64
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest

from outlook_cli.graph.mail import SendDraft, send_mail


def _draft() -> SendDraft:
    return SendDraft(
        to=["a@b.ca"],
        subject="hi",
        body="<p>hello</p>",
        body_html=True,
        cc=[],
        bcc=[],
        attachments=[],
        importance="normal",
        save_as_draft=False,
    )


def test_send_mail_posts_to_send_endpoint() -> None:
    client = MagicMock()
    client.post.return_value.status_code = 202
    send_mail(client, _draft())
    args, kwargs = client.post.call_args
    assert "/me/sendMail" in args[0]
    body = kwargs["json_body"]
    assert body["message"]["subject"] == "hi"
    assert body["message"]["toRecipients"][0]["emailAddress"]["address"] == "a@b.ca"
    assert body["message"]["body"]["contentType"] == "HTML"


def test_send_mail_text_body_when_not_html() -> None:
    client = MagicMock()
    draft = _draft()
    draft.body_html = False
    draft.body = "plain text"
    send_mail(client, draft)
    body = client.post.call_args.kwargs["json_body"]
    assert body["message"]["body"]["contentType"] == "Text"
    assert body["message"]["body"]["content"] == "plain text"


@pytest.mark.parametrize(
    ("api_base", "expected_type"),
    [
        (
            "https://outlook.office.com/api/v2.0",
            "#Microsoft.OutlookServices.FileAttachment",
        ),
        ("https://graph.microsoft.com/v1.0", "#microsoft.graph.fileAttachment"),
    ],
)
def test_send_mail_inlines_small_attachment(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    api_base: str,
    expected_type: str,
) -> None:
    """The attachment type must match the API being addressed.

    Sending Graph's name to outlook.office.com is rejected with
    400 UnableToDeserializePostBody.
    """
    monkeypatch.setenv("OUTLOOK_CLI_API_BASE", api_base)
    f = tmp_path / "small.txt"
    f.write_bytes(b"hello")
    client = MagicMock()
    draft = _draft()
    draft.attachments = [f]
    send_mail(client, draft)
    body = client.post.call_args.kwargs["json_body"]
    atts = body["message"]["attachments"]
    assert atts[0]["@odata.type"] == expected_type
    assert atts[0]["name"] == "small.txt"
    assert base64.b64decode(atts[0]["contentBytes"]) == b"hello"


def test_save_as_draft_uses_messages_endpoint() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "DRAFT-1"}
    draft = _draft()
    draft.save_as_draft = True
    send_mail(client, draft)
    assert any("/me/messages" in c.args[0] for c in client.post.call_args_list)


def test_large_attachment_uses_upload_session(tmp_path: Path) -> None:
    f = tmp_path / "big.bin"
    f.write_bytes(b"X" * (4 * 1024 * 1024))  # 4 MB > 3 MB threshold
    client = MagicMock()
    client.post.side_effect = [
        MagicMock(json=MagicMock(return_value={"id": "DRAFT-1"})),
        MagicMock(json=MagicMock(return_value={"uploadUrl": "https://upload.test/sess"})),
    ]
    draft = _draft()
    draft.attachments = [f]
    draft.save_as_draft = True
    with patch("outlook_cli.graph.mail.httpx.Client") as mock_client_cls:
        mock_raw = MagicMock()
        mock_client_cls.return_value.__enter__.return_value = mock_raw
        mock_raw.put.return_value.status_code = 200
        send_mail(client, draft)
    second_call = client.post.call_args_list[1]
    assert "/createUploadSession" in second_call.args[0]
    mock_raw.put.assert_called()


def test_large_attachment_raises_on_put_failure(tmp_path: Path) -> None:
    """Upload session PUT returning 4xx must raise GraphError, not be silently ignored."""
    from outlook_cli.graph.client import GraphError

    f = tmp_path / "big.bin"
    f.write_bytes(b"X" * (4 * 1024 * 1024))
    client = MagicMock()
    client.post.side_effect = [
        MagicMock(json=MagicMock(return_value={"id": "DRAFT-1"})),
        MagicMock(json=MagicMock(return_value={"uploadUrl": "https://upload.test/sess"})),
    ]
    draft = _draft()
    draft.attachments = [f]
    draft.save_as_draft = True
    with patch("outlook_cli.graph.mail.httpx.Client") as mock_client_cls:
        mock_raw = MagicMock()
        mock_client_cls.return_value.__enter__.return_value = mock_raw
        mock_raw.put.return_value.status_code = 413  # Payload Too Large
        mock_raw.put.return_value.text = "chunk too large"
        with pytest.raises(GraphError) as exc_info:
            send_mail(client, draft)
        assert exc_info.value.status_code == 413
