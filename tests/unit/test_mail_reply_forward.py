from unittest.mock import MagicMock

from outlook_cli.graph.mail import SendDraft, forward_mail, reply_mail


def _draft(reply_to: str = "MSG-1") -> SendDraft:
    return SendDraft(
        to=[],
        subject="Re: hello",
        body="my reply",
        body_html=False,
        cc=[],
        bcc=[],
        attachments=[],
        importance="normal",
        save_as_draft=False,
        reply_to_id=reply_to,
    )


def test_reply_posts_to_createReply_then_sends() -> None:  # noqa: N802
    client = MagicMock()
    client.post.side_effect = [
        MagicMock(json=MagicMock(return_value={"id": "DRAFT-X"})),  # createReply
        MagicMock(),  # send
    ]
    reply_mail(client, _draft(), reply_all=False)
    paths = [c.args[0] for c in client.post.call_args_list]
    assert "/me/messages/MSG-1/createReply" in paths
    assert any("/send" in p for p in paths)


def test_reply_all_uses_createReplyAll() -> None:  # noqa: N802
    client = MagicMock()
    client.post.side_effect = [
        MagicMock(json=MagicMock(return_value={"id": "DRAFT-X"})),
        MagicMock(),
    ]
    reply_mail(client, _draft(), reply_all=True)
    assert "/me/messages/MSG-1/createReplyAll" in client.post.call_args_list[0].args[0]


def test_forward_uses_createForward_with_to_recipients() -> None:  # noqa: N802
    client = MagicMock()
    client.post.side_effect = [
        MagicMock(json=MagicMock(return_value={"id": "DRAFT-X"})),
        MagicMock(),
    ]
    d = _draft()
    d.to = ["bob@example.com"]
    forward_mail(client, d)
    paths = [c.args[0] for c in client.post.call_args_list]
    assert "/me/messages/MSG-1/createForward" in paths
