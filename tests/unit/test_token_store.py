import os
import stat
from pathlib import Path

import pytest

from outlook_cli.auth.token_store import Credentials, load, save
from outlook_cli.errors import SessionExpired


def _sample_creds() -> Credentials:
    return Credentials(
        version=1,
        acquired_at="2026-05-22T08:37:14Z",
        tenant_id="11111111-2222-3333-4444-555555555555",
        client_id="9199bf20-a13f-4107-85dc-02114787ef48",
        home_account_id="00000000.11111111",
        username="pouriamortezaagha7@gmail.com",
        refresh_token="1.AXYA.SAMPLE",
        id_token_claims={"name": "Pouria", "preferred_username": "pouriamortezaagha7@gmail.com"},
    )


def test_save_then_load_round_trips(tmp_config_home: Path) -> None:
    creds = _sample_creds()
    save(creds)
    loaded = load()
    assert loaded == creds


def test_save_writes_0600_on_posix(tmp_config_home: Path) -> None:
    if os.name == "nt":
        pytest.skip("POSIX permissions not enforced on Windows")
    save(_sample_creds())
    mode = stat.S_IMODE(os.stat(tmp_config_home / "credentials.json").st_mode)
    assert mode == 0o600


def test_save_is_atomic_via_tmp_file(tmp_config_home: Path) -> None:
    save(_sample_creds())
    assert not (tmp_config_home / "credentials.json.tmp").exists()


def test_load_raises_session_expired_when_missing(tmp_config_home: Path) -> None:
    with pytest.raises(SessionExpired) as exc:
        load()
    assert "outlook login" in str(exc.value)


def test_load_raises_on_malformed_json(tmp_config_home: Path) -> None:
    (tmp_config_home / "credentials.json").write_text("not json")
    with pytest.raises(SessionExpired):
        load()


def test_rotate_refresh_token_replaces_atomically(tmp_config_home: Path) -> None:
    from outlook_cli.auth.token_store import rotate_refresh_token

    save(_sample_creds())
    rotate_refresh_token("1.AXYA.ROTATED")
    loaded = load()
    assert loaded.refresh_token == "1.AXYA.ROTATED"


def test_delete_removes_credentials_file(tmp_config_home: Path) -> None:
    from outlook_cli.auth.token_store import delete

    save(_sample_creds())
    delete()
    assert not (tmp_config_home / "credentials.json").exists()


def test_delete_is_idempotent(tmp_config_home: Path) -> None:
    from outlook_cli.auth.token_store import delete

    delete()
    delete()
