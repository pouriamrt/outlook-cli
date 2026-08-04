"""outlook login / logout / whoami commands."""

from __future__ import annotations

import json

import typer
from rich.console import Console

from outlook_cli.auth.login import capture_session_via_bookmarklet, save_session
from outlook_cli.auth.token_store import delete, load
from outlook_cli.errors import SessionExpired

err_console = Console(stderr=True)


def login() -> None:
    """Capture a Microsoft 365 session via browser bookmarklet.

    Prints a sign-in URL and a one-line JavaScript bookmarklet. You sign in in
    your normal browser, click the bookmarklet, and it sends your session to
    the CLI's temporary local HTTP server.

    No browser automation, no DevTools -- works with any tenant's
    SSO (federated, smart card, conditional access, whatever).
    """
    session = capture_session_via_bookmarklet()
    creds = save_session(session)
    typer.echo(f"Logged in as {creds.username}")
    if not session.is_foci:
        err_console.print(
            "[yellow]Warning:[/] refresh token is not marked as a Family RT (foci). "
            "Some scopes may require re-consent.",
        )


def logout() -> None:
    """Delete stored credentials. Idempotent."""
    delete()
    typer.echo("Logged out.")


def whoami(ctx: typer.Context) -> None:
    """Print signed-in user info."""
    try:
        creds = load()
    except SessionExpired as exc:
        err_console.print(str(exc))
        raise typer.Exit(code=77) from exc

    if ctx.obj and ctx.obj.get("json"):
        typer.echo(
            json.dumps(
                {
                    "username": creds.username,
                    "tenant_id": creds.tenant_id,
                    "client_id": creds.client_id,
                    "home_account_id": creds.home_account_id,
                    "id_token_claims": creds.id_token_claims,
                },
                indent=2,
            )
        )
    else:
        name = creds.id_token_claims.get("name", creds.username)
        typer.echo(f"{name} <{creds.username}>")
        typer.echo(f"tenant: {creds.tenant_id}")
