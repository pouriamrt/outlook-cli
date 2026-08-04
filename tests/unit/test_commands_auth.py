import json
from pathlib import Path
from unittest.mock import patch

from typer.testing import CliRunner

from outlook_cli.auth.login import ParsedSession
from outlook_cli.auth.token_store import Credentials, save
from outlook_cli.cli import app

runner = CliRunner()


def _sample_creds() -> Credentials:
    return Credentials(
        version=1,
        acquired_at="2026-05-22T08:37:14Z",
        tenant_id="t1",
        client_id="c1",
        home_account_id="a.t",
        username="pouriamortezaagha7@gmail.com",
        refresh_token="rt",
        id_token_claims={"name": "Pouria", "preferred_username": "pouriamortezaagha7@gmail.com"},
    )


def test_whoami_prints_username_when_logged_in(tmp_config_home: Path) -> None:
    save(_sample_creds())
    result = runner.invoke(app, ["whoami"])
    assert result.exit_code == 0
    assert "pouriamortezaagha7@gmail.com" in result.stdout


def test_whoami_json_emits_full_account(tmp_config_home: Path) -> None:
    save(_sample_creds())
    result = runner.invoke(app, ["--json", "whoami"])
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["username"] == "pouriamortezaagha7@gmail.com"
    assert payload["tenant_id"] == "t1"


def test_whoami_exits_77_when_logged_out(tmp_config_home: Path) -> None:
    result = runner.invoke(app, ["whoami"])
    assert result.exit_code == 77


def test_logout_removes_credentials(tmp_config_home: Path) -> None:
    save(_sample_creds())
    result = runner.invoke(app, ["logout"])
    assert result.exit_code == 0
    assert not (tmp_config_home / "credentials.json").exists()


def test_logout_is_idempotent(tmp_config_home: Path) -> None:
    result = runner.invoke(app, ["logout"])
    assert result.exit_code == 0


def test_login_invokes_bookmarklet_capture(tmp_config_home: Path) -> None:
    session = ParsedSession(
        refresh_token="rt-new",
        client_id="c1",
        tenant_id="t1",
        home_account_id="a.t",
        username="pouriamortezaagha7@gmail.com",
        id_token_claims={"name": "Pouria"},
        is_foci=True,
    )
    with patch(
        "outlook_cli.commands.auth.capture_session_via_bookmarklet",
        return_value=session,
    ) as mock_capture:
        result = runner.invoke(app, ["login"])
    assert result.exit_code == 0
    mock_capture.assert_called_once()
    from outlook_cli.auth.token_store import load

    assert load().username == "pouriamortezaagha7@gmail.com"
