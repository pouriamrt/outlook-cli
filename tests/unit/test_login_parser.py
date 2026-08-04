import json
from pathlib import Path

import pytest

from outlook_cli.auth.login import parse_msal_localstorage
from outlook_cli.errors import UserError

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def _load_fixture() -> dict[str, str]:
    return json.loads((FIXTURES / "msal_localstorage.json").read_text())


def test_parse_extracts_refresh_token() -> None:
    parsed = parse_msal_localstorage(_load_fixture())
    assert parsed.refresh_token == "1.AXYA.SAMPLE-REFRESH-TOKEN"


def test_parse_extracts_tenant_and_client() -> None:
    parsed = parse_msal_localstorage(_load_fixture())
    assert parsed.tenant_id == "11111111-2222-3333-4444-555555555555"
    assert parsed.client_id == "9199bf20-a13f-4107-85dc-02114787ef48"


def test_parse_extracts_username_and_home_account() -> None:
    parsed = parse_msal_localstorage(_load_fixture())
    assert parsed.username == "pouriamortezaagha7@gmail.com"
    assert parsed.home_account_id.startswith("00000000-")


def test_parse_decodes_id_token_claims() -> None:
    parsed = parse_msal_localstorage(_load_fixture())
    assert parsed.id_token_claims["name"] == "Pouria Mortezaagha"
    assert parsed.id_token_claims["preferred_username"] == "pouriamortezaagha7@gmail.com"


def test_parse_records_foci() -> None:
    parsed = parse_msal_localstorage(_load_fixture())
    assert parsed.is_foci is True


def test_parse_handles_msaljs_v5_pipe_delimited_keys() -> None:
    """Real MSAL.js v5 uses pipe-delimited cache keys; parser must handle them."""
    storage = json.loads((FIXTURES / "msal_localstorage_v5.json").read_text())
    parsed = parse_msal_localstorage(storage)
    # Prefer the most recently issued refresh token (msal.3 over msal.2).
    assert parsed.refresh_token == "NEW-REFRESH-TOKEN-V3"
    assert parsed.username == "pouriamortezaagha7@gmail.com"
    assert parsed.client_id == "9199bf20-a13f-4107-85dc-02114787ef48"
    assert parsed.tenant_id == "11111111-2222-3333-4444-555555555555"
    assert parsed.id_token_claims["name"] == "Pouria Mortezaagha"


def test_parse_raises_userror_when_no_refresh_token() -> None:
    storage = {"some-unrelated-key": "value"}
    with pytest.raises(UserError) as exc:
        parse_msal_localstorage(storage)
    assert "refresh" in str(exc.value).lower()
