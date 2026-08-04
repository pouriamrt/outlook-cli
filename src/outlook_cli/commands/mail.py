"""outlook mail ... commands."""

from __future__ import annotations

import base64
import sys
from dataclasses import replace
from pathlib import Path as _Path

import click
import typer
from rich.console import Console

from outlook_cli.dates import parse_human, parse_since
from outlook_cli.errors import NotFound, UserError
from outlook_cli.graph.client import GraphClient
from outlook_cli.graph.folders import resolve_folder_id
from outlook_cli.graph.mail import (
    MailListFilters,
    SendDraft,
    delete_message,
    flag_message,
    forward_mail,
    list_messages,
    mark_message,
    move_message,
    read_message,
    read_thread,
    reply_mail,
    search_messages,
    send_mail,
    unflag_message,
)
from outlook_cli.index_cache import resolve as resolve_index
from outlook_cli.index_cache import store as store_index
from outlook_cli.render.detail import render_message_detail
from outlook_cli.render.json_out import dump_mail_detail, dump_mail_list
from outlook_cli.render.tables import render_mail_list

mail_app = typer.Typer(name="mail", help="Mail operations.", no_args_is_help=True)


def _default_folder() -> str:
    """Read default_folder from config, falling back to 'inbox' if not set or unreadable."""
    # Function-scope import: breaks circular dep with commands.meta
    from outlook_cli.commands.meta import get_config

    return str(get_config("default_folder", "inbox"))


@mail_app.command("list")
def list_cmd(
    ctx: typer.Context,
    folder: str | None = typer.Option(None, "--folder", help="Folder name or Graph ID."),
    unread: bool = typer.Option(False, "--unread"),
    flagged: bool = typer.Option(False, "--flagged"),
    from_addr: str | None = typer.Option(None, "--from"),
    subject: str | None = typer.Option(None, "--subject"),
    since: str | None = typer.Option(None, "--since", help='"2d", "yesterday", ISO datetime'),
    top: int = typer.Option(25, "--top"),
    skip: int = typer.Option(0, "--skip"),
    all_pages: bool = typer.Option(False, "--all", help="Follow @odata.nextLink"),
) -> None:
    """List messages in a folder."""
    if folder is None:
        folder = _default_folder()
    client = GraphClient()
    folder_id = resolve_folder_id(folder, client)
    since_dt = parse_since(since) if since else None
    filters = MailListFilters(
        unread=unread,
        flagged=flagged,
        from_addr=from_addr,
        subject=subject,
        since=since_dt,
        top=top,
        skip=skip,
    )
    collected = []
    result = list_messages(client, folder_id=folder_id, filters=filters)
    collected.extend(result.items)
    while all_pages and result.next_link:
        next_skip = (filters.skip or 0) + len(result.items)
        filters = replace(filters, skip=next_skip)
        result = list_messages(client, folder_id=folder_id, filters=filters)
        collected.extend(result.items)

    items_payload = [{"index": i, "id": m.id} for i, m in enumerate(collected, start=1)]
    store_index("mail", items_payload)

    if ctx.obj and ctx.obj.get("json"):
        next_skip_out = (filters.skip + len(result.items)) if result.next_link else None
        typer.echo(dump_mail_list(collected, next_skip=next_skip_out))
    else:
        render_mail_list(Console(), collected)


@mail_app.command("read")
def read_cmd(
    ctx: typer.Context,
    ref: str = typer.Argument(..., help="Short index from last 'mail list' or full Graph ID."),
    raw: bool = typer.Option(False, "--raw", help="Do not convert HTML body to markdown."),
    save_attachments: str | None = typer.Option(
        None, "--save-attachments", help="Directory to save attachments to."
    ),
) -> None:
    """Render a single message."""
    try:
        message_id = resolve_index("mail", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc
    client = GraphClient()
    msg = read_message(client, message_id=message_id)
    if save_attachments:
        _save_attachments(client, message_id, save_attachments)
    if ctx.obj and ctx.obj.get("json"):
        typer.echo(dump_mail_detail(msg))
    else:
        render_message_detail(Console(), msg, raw=raw)


def _save_attachments(client: GraphClient, message_id: str, directory: str) -> None:
    out_dir = _Path(directory).resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    resp = client.get(f"/me/messages/{message_id}/attachments")
    for att in resp.json().get("value", []):
        if not att.get("@odata.type", "").endswith("#microsoft.graph.fileAttachment"):
            continue
        # Strip any directory components from Graph-supplied filename
        raw_name = att.get("name") or ""
        safe_name = _Path(raw_name).name
        if not safe_name:  # name was empty or pure path components
            continue
        raw_bytes = att.get("contentBytes")
        if not raw_bytes:
            continue
        dest = (out_dir / safe_name).resolve()
        # Defense-in-depth: refuse to write outside out_dir
        try:
            dest.relative_to(out_dir)
        except ValueError as exc:
            raise UserError(
                f"Refusing to save attachment outside target dir: {raw_name!r}"
            ) from exc
        content = base64.b64decode(raw_bytes)
        dest.write_bytes(content)


@mail_app.command("thread")
def thread_cmd(
    ctx: typer.Context,
    ref: str = typer.Argument(..., help="Index/ID of any message in the thread."),
    raw: bool = typer.Option(False, "--raw"),
) -> None:
    """Render the whole conversation chronologically."""
    try:
        message_id = resolve_index("mail", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc
    client = GraphClient()
    anchor = read_message(client, message_id=message_id)
    messages = read_thread(client, conversation_id=anchor.conversation_id)
    if ctx.obj and ctx.obj.get("json"):
        typer.echo(dump_mail_list(messages, next_skip=None))
    else:
        console = Console()
        for i, m in enumerate(messages, start=1):
            console.rule(f"[{i}/{len(messages)}]")
            render_message_detail(console, m, raw=raw)


def _gather_body(body_flag: str | None) -> str:
    if body_flag == "-":
        return sys.stdin.read()
    if body_flag is not None:
        return body_flag
    edited = click.edit("\n# Type your message above. Save and close to send.\n")
    return (edited or "").split("# Type your message above")[0].rstrip()


@mail_app.command("send")
def send_cmd(
    ctx: typer.Context,
    to: list[str] = typer.Option(  # noqa: B008
        [], "--to", help="Recipient address (repeatable)."
    ),
    cc: list[str] = typer.Option([], "--cc"),  # noqa: B008
    bcc: list[str] = typer.Option([], "--bcc"),  # noqa: B008
    subject: str = typer.Option(..., "--subject"),
    body: str | None = typer.Option(
        None, "--body", help='Body text, "-" for stdin, omit for $EDITOR.'
    ),
    attach: list[_Path] = typer.Option(  # noqa: B008
        [], "--attach", help="Attachment file path (repeatable)."
    ),
    html: bool = typer.Option(False, "--html", help="Body is HTML."),
    draft: bool = typer.Option(False, "--draft", help="Save as draft, do not send."),
    importance: str = typer.Option("normal", "--importance"),
) -> None:
    """Compose and send mail. With no --body, opens $EDITOR."""
    if not to:
        Console(stderr=True).print("--to is required.")
        raise typer.Exit(code=2)
    body_text = _gather_body(body)
    if not body_text.strip():
        proceed = typer.confirm("Body is empty. Send anyway?", default=False)
        if not proceed:
            raise typer.Exit(code=1)
    sd = SendDraft(
        to=to,
        cc=cc,
        bcc=bcc,
        subject=subject,
        body=body_text,
        body_html=html,
        attachments=list(attach),
        importance=importance,
        save_as_draft=draft,
    )
    client = GraphClient()
    send_mail(client, sd)
    typer.echo("Draft saved." if draft else "Sent.")


@mail_app.command("reply")
def reply_cmd(
    ref: str = typer.Argument(..., help="Index/ID of the message to reply to."),
    body: str | None = typer.Option(None, "--body"),
    cc: list[str] = typer.Option([], "--cc"),  # noqa: B008
    bcc: list[str] = typer.Option([], "--bcc"),  # noqa: B008
    attach: list[_Path] = typer.Option([], "--attach"),  # noqa: B008
    html: bool = typer.Option(False, "--html"),
    all_recipients: bool = typer.Option(False, "--all", help="Reply-all."),
    draft_flag: bool = typer.Option(False, "--draft"),
) -> None:
    """Reply to a message."""
    try:
        message_id = resolve_index("mail", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc
    body_text = _gather_body(body)
    if not body_text.strip() and not typer.confirm("Body is empty. Send anyway?", default=False):
        raise typer.Exit(code=1)
    sd = SendDraft(
        to=[],
        cc=cc,
        bcc=bcc,
        subject="",
        body=body_text,
        body_html=html,
        attachments=list(attach),
        reply_to_id=message_id,
        save_as_draft=draft_flag,
    )
    client = GraphClient()
    reply_mail(client, sd, reply_all=all_recipients)
    typer.echo("Draft saved." if draft_flag else "Sent.")


@mail_app.command("forward")
def forward_cmd(
    ref: str = typer.Argument(..., help="Index/ID of the message to forward."),
    to: list[str] = typer.Option([], "--to"),  # noqa: B008
    body: str | None = typer.Option(None, "--body"),
    attach: list[_Path] = typer.Option([], "--attach"),  # noqa: B008
    html: bool = typer.Option(False, "--html"),
    draft_flag: bool = typer.Option(False, "--draft"),
) -> None:
    """Forward a message."""
    if not to:
        Console(stderr=True).print("--to is required for forward.")
        raise typer.Exit(code=2)
    try:
        message_id = resolve_index("mail", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc
    body_text = _gather_body(body)
    sd = SendDraft(
        to=to,
        cc=[],
        bcc=[],
        subject="",
        body=body_text,
        body_html=html,
        attachments=list(attach),
        reply_to_id=message_id,
        save_as_draft=draft_flag,
    )
    client = GraphClient()
    forward_mail(client, sd)
    typer.echo("Draft saved." if draft_flag else "Forwarded.")


def _resolve_or_exit(ref: str) -> str:
    try:
        return resolve_index("mail", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc


@mail_app.command("move")
def move_cmd(
    ref: str = typer.Argument(...),
    folder: str = typer.Argument(..., help="Destination folder name or Graph ID."),
) -> None:
    """Move a message to a folder."""
    message_id = _resolve_or_exit(ref)
    client = GraphClient()
    folder_id = resolve_folder_id(folder, client)
    move_message(client, message_id=message_id, destination_folder_id=folder_id)
    typer.echo("Moved.")


@mail_app.command("delete")
def delete_cmd(
    ref: str = typer.Argument(...),
    purge: bool = typer.Option(False, "--purge", help="Hard delete (not recoverable)."),
) -> None:
    """Move message to Deleted Items (or hard-delete with --purge)."""
    message_id = _resolve_or_exit(ref)
    client = GraphClient()
    delete_message(client, message_id=message_id, purge=purge)
    typer.echo("Purged." if purge else "Deleted.")


@mail_app.command("flag")
def flag_cmd(
    ref: str = typer.Argument(...),
    due: str | None = typer.Option(None, "--due", help='e.g. "fri", "2026-05-30"'),
) -> None:
    """Flag a message for follow-up."""
    message_id = _resolve_or_exit(ref)
    due_dt = parse_human(due) if due else None
    client = GraphClient()
    flag_message(client, message_id=message_id, due=due_dt)
    typer.echo("Flagged.")


@mail_app.command("unflag")
def unflag_cmd(ref: str = typer.Argument(...)) -> None:
    """Remove follow-up flag from a message."""
    message_id = _resolve_or_exit(ref)
    client = GraphClient()
    unflag_message(client, message_id=message_id)
    typer.echo("Unflagged.")


@mail_app.command("mark")
def mark_cmd(
    ref: str = typer.Argument(...),
    read: bool | None = typer.Option(None, "--read/--unread"),
) -> None:
    """Mark a message read or unread."""
    if read is None:
        Console(stderr=True).print("Specify --read or --unread.")
        raise typer.Exit(code=2)
    message_id = _resolve_or_exit(ref)
    client = GraphClient()
    mark_message(client, message_id=message_id, read=read)
    typer.echo("Marked.")


@mail_app.command("search")
def search_cmd(
    ctx: typer.Context,
    query: str = typer.Argument(..., help='KQL-style query, e.g. "from:alice subject:Q3"'),
    folder: str | None = typer.Option(None, "--folder"),
    top: int = typer.Option(25, "--top"),
) -> None:
    """KQL-style search across mailbox (or within a folder)."""
    client = GraphClient()
    folder_id = resolve_folder_id(folder, client) if folder else None
    messages = search_messages(client, query=query, folder_id=folder_id, top=top)
    store_index("mail", [{"index": i, "id": m.id} for i, m in enumerate(messages, start=1)])
    if ctx.obj and ctx.obj.get("json"):
        typer.echo(dump_mail_list(messages, next_skip=None))
    else:
        render_mail_list(Console(), messages)
