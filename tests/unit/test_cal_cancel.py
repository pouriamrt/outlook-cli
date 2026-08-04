from unittest.mock import MagicMock

from outlook_cli.graph.calendar import cancel_event


def test_cancel_event_posts_to_cancel_endpoint() -> None:
    client = MagicMock()
    cancel_event(client, event_id="EV-1", comment="Conflict")
    args, kwargs = client.post.call_args
    assert "/me/events/EV-1/cancel" in args[0]
    assert kwargs["json_body"]["Comment"] == "Conflict"


def test_cancel_event_default_comment_empty() -> None:
    client = MagicMock()
    cancel_event(client, event_id="EV-1", comment="")
    assert client.post.call_args.kwargs["json_body"]["Comment"] == ""
