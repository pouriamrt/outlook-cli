"""Detail-view renderers for a single Message or Event."""

from __future__ import annotations

import html2text
from rich.console import Console
from rich.markdown import Markdown
from rich.table import Table

from outlook_cli.graph.models import Event, Message


def _html_to_md(html: str) -> str:
    h = html2text.HTML2Text()
    h.body_width = 0
    h.ignore_images = True
    result: str = h.handle(html).strip()
    return result


def render_message_detail(console: Console, msg: Message, *, raw: bool) -> None:
    header = Table.grid(padding=(0, 1))
    header.add_column(style="bold cyan", justify="right")
    header.add_column()
    header.add_row("From", f"{msg.from_.name} <{msg.from_.address}>")
    if msg.to:
        header.add_row("To", ", ".join(f"{r.name} <{r.address}>" for r in msg.to))
    if msg.cc:
        header.add_row("Cc", ", ".join(f"{r.name} <{r.address}>" for r in msg.cc))
    header.add_row("Subject", msg.subject or "(no subject)")
    header.add_row("Date", msg.received_at.astimezone().strftime("%a %b %d %Y %H:%M"))
    if msg.importance != "normal":
        header.add_row("Importance", msg.importance)
    console.print(header)
    console.print()
    if raw or msg.body_content_type == "text":
        console.print(msg.body_html)
    else:
        console.print(Markdown(_html_to_md(msg.body_html)))


def render_event_detail(console: Console, ev: Event, *, show_attendees: bool) -> None:
    header = Table.grid(padding=(0, 1))
    header.add_column(style="bold cyan", justify="right")
    header.add_column()
    header.add_row("Subject", ev.subject or "(no subject)")
    header.add_row(
        "When",
        f"{ev.start.astimezone():%a %b %d %H:%M} → {ev.end.astimezone():%H:%M}",
    )
    if ev.location:
        header.add_row("Where", ev.location)
    if ev.is_online_meeting and ev.online_meeting_url:
        header.add_row("Teams", ev.online_meeting_url)
    header.add_row("Organizer", f"{ev.organizer.name} <{ev.organizer.address}>")
    console.print(header)
    if show_attendees and ev.attendees:
        atable = Table(title="Attendees", show_header=True)
        atable.add_column("Name")
        atable.add_column("Address")
        atable.add_column("Response")
        for a in ev.attendees:
            atable.add_row(a.name, a.address, a.response)
        console.print(atable)
    console.print()
    if ev.body_html:
        console.print(Markdown(_html_to_md(ev.body_html)))
