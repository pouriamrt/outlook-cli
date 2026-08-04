from pathlib import Path
from unittest.mock import MagicMock

import pytest
from freezegun import freeze_time

from outlook_cli.graph.folders import WELL_KNOWN, resolve_folder_id


def _fake_client() -> MagicMock:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {"id": "INBOX-ID", "displayName": "Inbox", "totalItemCount": 0, "unreadItemCount": 0},
            {
                "id": "PROJX-ID",
                "displayName": "ProjectX",
                "totalItemCount": 0,
                "unreadItemCount": 0,
            },
        ]
    }
    return client


def test_well_known_name_resolves_without_api_call(tmp_cache_home: Path) -> None:
    client = MagicMock()
    assert resolve_folder_id("inbox", client) == WELL_KNOWN["inbox"]
    client.get.assert_not_called()


def test_custom_folder_name_fetches_via_graph(tmp_cache_home: Path) -> None:
    client = _fake_client()
    assert resolve_folder_id("ProjectX", client) == "PROJX-ID"
    client.get.assert_called_once()


def test_custom_folder_cached_for_subsequent_calls(tmp_cache_home: Path) -> None:
    client = _fake_client()
    with freeze_time("2026-05-22T10:00:00Z"):
        resolve_folder_id("ProjectX", client)
        resolve_folder_id("ProjectX", client)
    assert client.get.call_count == 1


def test_cache_expires_after_24h(tmp_cache_home: Path) -> None:
    client = _fake_client()
    with freeze_time("2026-05-22T10:00:00Z"):
        resolve_folder_id("ProjectX", client)
    with freeze_time("2026-05-23T11:00:00Z"):
        resolve_folder_id("ProjectX", client)
    assert client.get.call_count == 2


def test_unknown_folder_raises(tmp_cache_home: Path) -> None:
    client = _fake_client()
    with pytest.raises(Exception, match="not found"):
        resolve_folder_id("Nonexistent", client)


def test_passthrough_for_full_graph_id(tmp_cache_home: Path) -> None:
    client = MagicMock()
    folder_id = "AAMkAGI2NzNkY2I5LWFiY2QtMTIzNA=="
    assert resolve_folder_id(folder_id, client) == folder_id
    client.get.assert_not_called()
