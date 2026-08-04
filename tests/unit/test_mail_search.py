from unittest.mock import MagicMock

from outlook_cli.graph.mail import search_messages


def test_search_passes_query_as_dollar_search() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": []}
    search_messages(client, query="from:alice subject:Q3", folder_id=None, top=25)
    args, kwargs = client.get.call_args
    assert "/me/messages" in args[0]
    assert kwargs["params"]["$search"] == '"from:alice subject:Q3"'


def test_search_in_folder_uses_folder_endpoint() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": []}
    search_messages(client, query="test", folder_id="inbox", top=10)
    args, _ = client.get.call_args
    assert "/me/mailFolders/inbox/messages" in args[0]


def test_search_returns_parsed_messages() -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {
                "id": "S1",
                "subject": "found",
                "receivedDateTime": "2026-05-22T10:00:00Z",
                "from": {"emailAddress": {"name": "A", "address": "a@b"}},
            }
        ]
    }
    msgs = search_messages(client, query="found", folder_id=None, top=5)
    assert msgs[0].subject == "found"
