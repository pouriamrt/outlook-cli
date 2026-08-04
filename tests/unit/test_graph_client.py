import respx
from httpx import Response

from outlook_cli.graph.client import GraphClient, GraphError


@respx.mock
def test_client_injects_bearer_token() -> None:
    route = respx.get("https://outlook.office.com/api/v2.0/me").mock(
        return_value=Response(200, json={"id": "x"})
    )
    client = GraphClient(token_provider=lambda scope: "AT-1")
    resp = client.get("/me")
    assert resp.status_code == 200
    req = route.calls.last.request
    assert req.headers["Authorization"] == "Bearer AT-1"


@respx.mock
def test_client_uses_base_url_for_relative_paths() -> None:
    respx.get("https://outlook.office.com/api/v2.0/me/messages").mock(
        return_value=Response(200, json={"value": []})
    )
    client = GraphClient(token_provider=lambda scope: "AT-1")
    resp = client.get("/me/messages")
    assert resp.status_code == 200


@respx.mock
def test_client_get_passes_query_params() -> None:
    route = respx.get("https://outlook.office.com/api/v2.0/me/messages").mock(
        return_value=Response(200, json={"value": []})
    )
    client = GraphClient(token_provider=lambda scope: "AT-1")
    client.get("/me/messages", params={"$top": 5, "$select": "id,subject"})
    req = route.calls.last.request
    assert "%24top=5" in str(req.url)
    assert "%24select=id%2Csubject" in str(req.url)


@respx.mock
def test_client_post_pascal_cases_for_outlook_rest() -> None:
    """Outlook REST v2.0 endpoints want PascalCase JSON keys; the client must denormalize."""
    import json

    respx.post("https://outlook.office.com/api/v2.0/me/sendMail").mock(return_value=Response(202))
    client = GraphClient(token_provider=lambda scope: "AT-1")
    client.post("/me/sendMail", json_body={"message": {"subject": "hi"}})
    req = respx.routes[0].calls.last.request
    assert req.headers["content-type"].startswith("application/json")
    assert json.loads(req.content) == {"Message": {"Subject": "hi"}}


@respx.mock
def test_client_post_preserves_camel_case_for_graph() -> None:
    """Microsoft Graph wants camelCase; when base_url is graph.microsoft.com, leave as-is."""
    import json

    respx.post("https://graph.microsoft.com/v1.0/me/sendMail").mock(return_value=Response(202))
    client = GraphClient(
        token_provider=lambda scope: "AT-1",
        base_url="https://graph.microsoft.com/v1.0",
    )
    client.post("/me/sendMail", json_body={"message": {"subject": "hi"}})
    req = respx.routes[0].calls.last.request
    assert json.loads(req.content) == {"message": {"subject": "hi"}}


@respx.mock
def test_client_post_preserves_odata_metadata_keys() -> None:
    """@odata.type / @odata.bind are OData metadata and must not be PascalCased."""
    import json

    respx.post("https://outlook.office.com/api/v2.0/me/messages").mock(
        return_value=Response(201, json={"id": "X"})
    )
    client = GraphClient(token_provider=lambda scope: "AT-1")
    client.post(
        "/me/messages",
        json_body={
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": "x.txt",
                    "contentBytes": "AA==",
                }
            ],
        },
    )
    req = respx.routes[0].calls.last.request
    sent = json.loads(req.content)
    assert sent["Attachments"][0]["@odata.type"] == "#microsoft.graph.fileAttachment"
    assert sent["Attachments"][0]["Name"] == "x.txt"
    assert sent["Attachments"][0]["ContentBytes"] == "AA=="


@respx.mock
def test_client_raises_for_4xx() -> None:
    respx.get("https://outlook.office.com/api/v2.0/me/messages/bad").mock(
        return_value=Response(404, json={"error": {"code": "ItemNotFound"}})
    )
    client = GraphClient(token_provider=lambda scope: "AT-1")
    import pytest

    with pytest.raises(GraphError) as exc:
        client.get("/me/messages/bad")
    assert exc.value.status_code == 404
