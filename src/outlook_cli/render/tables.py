"""Rich tables for list-style output."""

from __future__ import annotations

from datetime import datetime

from rich.console import Console
from rich.table import Table
from rich.text import Text

from outlook_cli.graph.models import Event, Message


def _fmt_date(dt: datetime) -> str:
    return dt.astimezone().strftime("%a %b %d %H:%M")


def render_mail_list(console: Console, messages: list[Message]) -> None:
    if not messages:
        console.print("[dim]No messages.[/dim]")
        return
    table = Table(show_header=True, header_style="bold cyan", box=None, expand=True)
    table.add_column("#", width=3, justify="right")
    table.add_column("Flags", width=3)
    table.add_column("From", width=24, overflow="ellipsis")
    table.add_column("Subject", overflow="ellipsis")
    table.add_column("Received", width=18)
    for idx, msg in enumerate(messages, start=1):
        flags = ""
        if not msg.is_read:
            flags += "●"
        if msg.is_flagged:
            flags += "🚩"
        if msg.has_attachments:
            flags += "📎"
        from_text = Text(msg.from_.name or msg.from_.address or "—")
        subj_text = Text(msg.subject or "(no subject)")
        if not msg.is_read:
            from_text.stylize("bold")
            subj_text.stylize("bold")
        table.add_row(str(idx), flags, from_text, subj_text, _fmt_date(msg.received_at))
    console.print(table)


def render_event_list(console: Console, events: list[Event]) -> None:
    if not events:
        console.print("[dim]No events.[/dim]")
        return
    table = Table(show_header=True, header_style="bold cyan", box=None, expand=True)
    table.add_column("#", width=3, justify="right")
    table.add_column("When", width=22)
    table.add_column("Subject", overflow="ellipsis")
    table.add_column("Where", width=20, overflow="ellipsis")
    for idx, ev in enumerate(events, start=1):
        when = f"{_fmt_date(ev.start)} ({_duration(ev.start, ev.end)})"
        where = "Teams" if ev.is_online_meeting else (ev.location or "—")
        table.add_row(str(idx), when, ev.subject or "(no subject)", where)
    console.print(table)


def _duration(start: datetime, end: datetime) -> str:
    minutes = int((end - start).total_seconds() // 60)
    if minutes < 60:
        return f"{minutes}m"
    return f"{minutes // 60}h{minutes % 60:02d}m" if minutes % 60 else f"{minutes // 60}h"
