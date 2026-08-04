"""outlook config / version / --json-schema commands."""

from __future__ import annotations

import tomllib
from pathlib import Path
from typing import Any

import tomli_w
import typer
from rich.console import Console

from outlook_cli.config import config_home, ensure_dirs

config_app = typer.Typer(name="config", help="Show/set CLI config.", no_args_is_help=True)

_DEFAULTS: dict[str, Any] = {
    "default_folder": "inbox",
    "default_signature_file": "",
    "reply_quote_style": "outlook",
    "date_format": "short",
    "table_max_width": 0,
    "editor": "",
}


def _config_path() -> Path:
    return config_home() / "config.toml"


def _load() -> dict[str, Any]:
    p = _config_path()
    if not p.exists():
        return dict(_DEFAULTS)
    with p.open("rb") as f:
        data = tomllib.load(f)
    return {**_DEFAULTS, **data}


def _save(values: dict[str, Any]) -> None:
    ensure_dirs()
    p = _config_path()
    with p.open("wb") as f:
        tomli_w.dump(values, f)


def get_config(key: str, default: Any) -> Any:
    """Public accessor: returns config[key] or default on any read error.

    `tomllib.TOMLDecodeError` is a subclass of `ValueError`, so the catch
    also handles malformed TOML files.
    """
    try:
        return _load().get(key, default)
    except (OSError, ValueError) as exc:
        import logging

        logging.getLogger(__name__).debug("Config read failed for %r: %s", key, exc)
        return default


@config_app.command("get")
def get_cmd(key: str = typer.Argument(...)) -> None:
    values = _load()
    if key not in values:
        Console(stderr=True).print(f"Unknown config key: '{key}'.")
        raise typer.Exit(code=64)
    typer.echo(values[key])


@config_app.command("set")
def set_cmd(key: str = typer.Argument(...), value: str = typer.Argument(...)) -> None:
    if key not in _DEFAULTS:
        Console(stderr=True).print(f"Unknown config key: '{key}'.")
        raise typer.Exit(code=64)
    values = _load()
    default = _DEFAULTS[key]
    if isinstance(default, bool):
        values[key] = value.lower() in ("1", "true", "yes")
    elif isinstance(default, int):
        values[key] = int(value)
    else:
        values[key] = value
    _save(values)
    typer.echo(f"Set {key}={values[key]}")


@config_app.command("list")
def list_cmd() -> None:
    values = _load()
    for k, v in values.items():
        typer.echo(f"{k} = {v!r}")
