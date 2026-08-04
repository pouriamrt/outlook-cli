import time

import httpx
import pytest
import respx
from httpx import Response

from outlook_cli.graph.client import GraphClient, GraphError


@respx.mock
def test_429_retries_then_succeeds(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    route = respx.get("https://outlook.office.com/api/v2.0/me").mock(
        side_effect=[
            Response(429, headers={"Retry-After": "1"}, json={"error": {"code": "Throttled"}}),
            Response(429, headers={"Retry-After": "1"}, json={"error": {"code": "Throttled"}}),
            Response(200, json={"id": "x"}),
        ]
    )
    client = GraphClient(token_provider=lambda s: "AT")
    resp = client.get("/me")
    assert resp.status_code == 200
    assert route.call_count == 3


@respx.mock
def test_5xx_retries_with_backoff(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    respx.get("https://outlook.office.com/api/v2.0/me").mock(
        side_effect=[Response(503), Response(502), Response(200, json={"id": "x"})]
    )
    client = GraphClient(token_provider=lambda s: "AT")
    resp = client.get("/me")
    assert resp.status_code == 200


@respx.mock
def test_429_gives_up_after_max_attempts(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    respx.get("https://outlook.office.com/api/v2.0/me").mock(
        return_value=Response(429, headers={"Retry-After": "0"})
    )
    client = GraphClient(token_provider=lambda s: "AT")
    with pytest.raises(GraphError) as exc:
        client.get("/me")
    assert exc.value.status_code == 429


@respx.mock
def test_401_triggers_one_force_refresh_then_retries(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    respx.get("https://outlook.office.com/api/v2.0/me").mock(
        side_effect=[Response(401), Response(200, json={"id": "x"})]
    )
    tokens = iter(["AT-stale", "AT-fresh"])
    client = GraphClient(token_provider=lambda s: next(tokens))
    resp = client.get("/me")
    assert resp.status_code == 200


@respx.mock
def test_401_after_force_refresh_raises(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    respx.get("https://outlook.office.com/api/v2.0/me").mock(return_value=Response(401))
    client = GraphClient(token_provider=lambda s: "AT")
    with pytest.raises(GraphError) as exc:
        client.get("/me")
    assert exc.value.status_code == 401


@respx.mock
def test_network_error_retries(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(time, "sleep", lambda s: None)
    respx.get("https://outlook.office.com/api/v2.0/me").mock(
        side_effect=[httpx.ConnectError("network"), Response(200, json={"id": "x"})]
    )
    client = GraphClient(token_provider=lambda s: "AT")
    resp = client.get("/me")
    assert resp.status_code == 200
