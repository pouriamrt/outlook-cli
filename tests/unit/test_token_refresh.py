from pathlib import Path

import pytest
import respx
from freezegun import freeze_time
from httpx import Response

from outlook_cli.auth.token_refresh import GRAPH_SCOPE, get_token
from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.errors import SessionExpired


def _seed(tmp_config_home: Path) -> Credentials:
    creds = Credentials(
        version=1,
        acquired_at="2026-05-22T08:00:00Z",
        tenant_id="t1",
        client_id="c1",
        home_account_id="a1.t1",
        username="u@example.com",
        refresh_token="rt-original",
    )
    save(creds)
    return creds


@respx.mock
@freeze_time("2026-05-22T10:00:00Z")
def test_get_token_mints_access_token_when_cache_empty(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(
            200,
            json={
                "access_token": "at-1",
                "expires_in": 3599,
                "refresh_token": "rt-rotated",
                "token_type": "Bearer",
            },
        )
    )
    token = get_token(GRAPH_SCOPE)
    assert token == "at-1"


@respx.mock
@freeze_time("2026-05-22T10:00:00Z")
def test_get_token_rotates_refresh_token_when_present(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(
            200,
            json={"access_token": "at-1", "expires_in": 3599, "refresh_token": "rt-new"},
        )
    )
    get_token(GRAPH_SCOPE)
    from outlook_cli.auth.token_store import load

    assert load().refresh_token == "rt-new"


@respx.mock
@freeze_time("2026-05-22T10:00:00Z")
def test_get_token_leaves_rt_unchanged_when_response_omits_it(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(200, json={"access_token": "at-1", "expires_in": 3599})
    )
    get_token(GRAPH_SCOPE)
    from outlook_cli.auth.token_store import load

    assert load().refresh_token == "rt-original"


@respx.mock
def test_get_token_uses_cached_at_when_still_valid(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    mint = respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(200, json={"access_token": "at-cached", "expires_in": 3599})
    )
    with freeze_time("2026-05-22T10:00:00Z"):
        first = get_token(GRAPH_SCOPE)
    with freeze_time("2026-05-22T10:30:00Z"):
        second = get_token(GRAPH_SCOPE)
    assert first == second == "at-cached"
    assert mint.call_count == 1


@respx.mock
def test_get_token_refreshes_when_cache_expired(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        side_effect=[
            Response(200, json={"access_token": "at-1", "expires_in": 3599}),
            Response(200, json={"access_token": "at-2", "expires_in": 3599}),
        ]
    )
    with freeze_time("2026-05-22T10:00:00Z"):
        assert get_token(GRAPH_SCOPE) == "at-1"
    with freeze_time("2026-05-22T11:00:00Z"):
        assert get_token(GRAPH_SCOPE) == "at-2"


@respx.mock
@freeze_time("2026-05-22T10:00:00Z")
def test_get_token_raises_session_expired_on_invalid_grant(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(
            400, json={"error": "invalid_grant", "error_description": "RT expired"}
        )
    )
    with pytest.raises(SessionExpired):
        get_token(GRAPH_SCOPE)


@respx.mock
@freeze_time("2026-05-22T10:00:00Z")
def test_get_token_raises_session_expired_on_interaction_required(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        return_value=Response(400, json={"error": "interaction_required"})
    )
    with pytest.raises(SessionExpired):
        get_token(GRAPH_SCOPE)


@respx.mock
def test_separate_scopes_get_separate_cached_tokens(
    tmp_config_home: Path, tmp_cache_home: Path
) -> None:
    _seed(tmp_config_home)
    respx.post("https://login.microsoftonline.com/t1/oauth2/v2.0/token").mock(
        side_effect=[
            Response(200, json={"access_token": "at-graph", "expires_in": 3599}),
            Response(200, json={"access_token": "at-other", "expires_in": 3599}),
        ]
    )
    with freeze_time("2026-05-22T10:00:00Z"):
        assert get_token("https://graph.microsoft.com/.default") == "at-graph"
        assert get_token("https://outlook.office365.com/.default") == "at-other"
