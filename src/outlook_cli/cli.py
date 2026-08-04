"""Typer root app for outlook-cli."""

from __future__ import annotations

import json as _json

import typer

from outlook_cli.commands import auth as auth_cmd
from outlook_cli.commands.calendar import cal_app
from outlook_cli.commands.mail import mail_app
from outlook_cli.commands.meta import config_app
from outlook_cli.schemas import get_schema

app = typer.Typer(
    name="outlook",
    help="Microsoft 365 mail and calendar from the terminal.",
    no_args_is_help=True,
    rich_markup_mode="rich",
)


@app.callback(invoke_without_command=True)
def _root(
    ctx: typer.Context,
    json_output: bool = typer.Option(
        False, "--json", help="Emit machine-readable JSON instead of rich tables."
    ),
    json_schema: str | None = typer.Option(
        None,
        "--json-schema",
        help="Print JSON Schema for a command (e.g. 'mail.list').",
    ),
    verbose: int = typer.Option(
        0, "--verbose", "-v", count=True, help="Increase verbosity (-v, -vv)."
    ),
) -> None:
    """Microsoft 365 mail and calendar from the terminal."""
    if json_schema is not None:
        schema = get_schema(json_schema)
        if schema is None:
            typer.secho(f"Unknown schema: '{json_schema}'.", err=True, fg="red")
            raise typer.Exit(code=64)
        typer.echo(_json.dumps(schema, indent=2))
        raise typer.Exit(code=0)
    ctx.obj = {"json": json_output, "verbose": verbose}


@app.command()
def version() -> None:
    """Print CLI version + Graph API target."""
    from outlook_cli import __version__

    typer.echo(f"outlook-cli {__version__} (Microsoft Graph v1.0)")


app.command(name="login")(auth_cmd.login)
app.command(name="logout")(auth_cmd.logout)
app.command(name="whoami")(auth_cmd.whoami)

app.add_typer(mail_app, name="mail")
app.add_typer(cal_app, name="cal")
app.add_typer(config_app, name="config")


def main() -> None:
    """Entry point with SessionExpired handling (exits 77).

    The function-scope imports avoid a circular import between cli and errors
    that would otherwise be triggered at module load time.
    """
    import sys

    import typer as _typer

    from outlook_cli.errors import SessionExpired

    try:
        app()
    except SessionExpired as exc:
        _typer.secho(str(exc), err=True, fg=_typer.colors.RED)
        sys.exit(77)
