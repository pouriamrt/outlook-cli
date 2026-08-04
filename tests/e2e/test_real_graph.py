"""End-to-end smoke tests against real Microsoft Graph.

Gated by environment variable OUTLOOK_CLI_E2E=1. Requires a valid credentials.json
(run `outlook login` first).
"""

from __future__ import annotations

import json
import os
import uuid
from datetime import UTC, datetime, timedelta

import pytest
from typer.testing import CliRunner

from outlook_cli.cli import app

pytestmark = pytest.mark.skipif(
    os.environ.get("OUTLOOK_CLI_E2E") != "1",
    reason="E2E suite gated by OUTLOOK_CLI_E2E=1",
)

runner = CliRunner()


def test_e2e_whoami_succeeds() -> None:
    result = runner.invoke(app, ["whoami"])
    assert result.exit_code == 0
    assert "@" in result.stdout


def test_e2e_mail_list_inbox_returns_at_least_one() -> None:
    result = runner.invoke(app, ["--json", "mail", "list", "--top", "1"])
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert "items" in payload


def test_e2e_mail_search_does_not_crash() -> None:
    result = runner.invoke(app, ["--json", "mail", "search", "test", "--top", "1"])
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert "items" in payload


def test_e2e_cal_today_does_not_crash() -> None:
    result = runner.invoke(app, ["--json", "cal", "today"])
    assert result.exit_code == 0


def test_e2e_create_then_cancel_event() -> None:
    title = f"outlook-cli E2E {uuid.uuid4().hex[:8]}"
    start = (datetime.now(UTC) + timedelta(days=1)).strftime("%Y-%m-%dT15:00:00")
    create_result = runner.invoke(
        app,
        ["cal", "create", "--title", title, "--start", start, "--duration", "30m"],
    )
    assert create_result.exit_code == 0
    list_result = runner.invoke(app, ["--json", "cal", "week"])
    payload = json.loads(list_result.stdout)
    match = next((it for it in payload["items"] if it["subject"] == title), None)
    assert match, f"Created event '{title}' not found in calendar."
    idx = match["index"]
    cancel_result = runner.invoke(app, ["cal", "cancel", str(idx), "--comment", "E2E teardown"])
    assert cancel_result.exit_code == 0
