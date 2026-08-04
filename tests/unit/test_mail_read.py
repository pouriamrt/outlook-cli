import json
from pathlib import Path
from unittest.mock import MagicMock

from outlook_cli.graph.mail import read_message

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_read_message_calls_graph_with_id_and_full_select() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = json.loads(
        (FIXTURES / "graph_message.json").read_text()
    )
    msg = read_message(client, message_id="MSG-123")
    args, kwargs = client.get.call_args
    assert "/me/messages/MSG-123" in args[0]
    assert "body" in kwargs["params"]["$select"]
    assert msg.subject == "Re: Q3 planning"
    assert msg.body_html.startswith("<p>")
