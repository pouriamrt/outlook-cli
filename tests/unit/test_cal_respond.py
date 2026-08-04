from unittest.mock import MagicMock

import pytest

from outlook_cli.graph.calendar import respond_event


@pytest.mark.parametrize(
    "verb,endpoint",
    [
        ("accept", "accept"),
        ("decline", "decline"),
        ("tentative", "tentativelyAccept"),
    ],
)
def test_respond_event_posts_to_correct_endpoint(verb: str, endpoint: str) -> None:
    client = MagicMock()
    respond_event(client, event_id="EV-1", verb=verb, comment="Sure")
    args, kwargs = client.post.call_args
    assert f"/me/events/EV-1/{endpoint}" in args[0]
    assert kwargs["json_body"]["comment"] == "Sure"
    assert kwargs["json_body"]["sendResponse"] is True


def test_respond_event_raises_on_unknown_verb() -> None:
    client = MagicMock()
    with pytest.raises(ValueError, match="Unknown response verb"):
        respond_event(client, event_id="EV", verb="maybe", comment="")
