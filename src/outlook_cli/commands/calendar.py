"""outlook cal ... commands."""

from __future__ import annotations

import json
import re
from datetime import date, datetime, time, timedelta
from typing import TypeVar

import click
import typer
from rich.console import Console

from outlook_cli.dates import parse_human
from outlook_cli.errors import NotFound
from outlook_cli.graph.calendar import (
    CreateEvent,
    cancel_event,
    create_event,
    event_to_ics,
    find_meeting_times,
    get_event,
    list_events,
    respond_event,
)
from outlook_cli.graph.client import GraphClient
from outlook_cli.graph.models import Event, FindTimeSuggestion
from outlook_cli.index_cache import resolve as resolve_index
from outlook_cli.index_cache import store as store_index
from outlook_cli.render.detail import render_event_detail
from outlook_cli.render.json_out import dump_event_detail, dump_event_list
from outlook_cli.render.tables import render_event_list

cal_app = typer.Typer(name="cal", help="Calendar operations.", no_args_is_help=True)


def _local_midnight(day_offset: int = 0) -> datetime:
    today = datetime.now().astimezone().date()
    target = today + timedelta(days=day_offset)
    return datetime.combine(target, time.min).astimezone()


# Working hours used by find-time when probing each day for free slots.
_WORK_START = time(9, 0)
_WORK_END = time(17, 0)


def _business_days(window: str, *, today: date) -> list[date]:
    """Resolve a --window keyword to the dates find-time should probe.

    'next week'  -> next calendar week's Monday..Friday.
    'this week' / 'week' -> this calendar week's Monday..Friday, dropping days
                    already past; on a weekend (none remain) it rolls to next week.
    anything else -> a single day parsed from the text (e.g. "tomorrow",
                    "next monday", a date). Matching is exact so phrases like
                    "next monday" resolve to that one day, not a whole week.
    """
    keyword = window.strip().lower()
    if keyword == "next week":
        monday = today - timedelta(days=today.weekday()) + timedelta(days=7)
        return [monday + timedelta(days=i) for i in range(5)]
    if keyword in ("this week", "week"):
        monday = today - timedelta(days=today.weekday())
        week = [monday + timedelta(days=i) for i in range(5)]
        days = [d for d in week if d >= today]
        return days or _business_days("next week", today=today)
    return [parse_human(window).date()]


def _day_window(day: date, *, now: datetime) -> tuple[datetime, datetime] | None:
    """Working-hours window (local 09:00-17:00) for ``day``, or None if fully past.

    For the current day the start is clamped to ``now`` so no past slots are offered.
    """
    start = datetime.combine(day, _WORK_START).astimezone()
    end = datetime.combine(day, _WORK_END).astimezone()
    if end <= now:
        return None
    return (max(start, now), end)


# Per-day, fetch the full set of free slots from Graph, then sample --per-day
# across it. High enough to cover a working day at fine granularity; Graph
# returns fewer when fewer slots are free.
_DAY_FETCH_CANDIDATES = 100

_T = TypeVar("_T")


def _spread(items: list[_T], k: int) -> list[_T]:
    """Evenly sample up to ``k`` items across ``items`` (chronological order).

    Graph returns free slots earliest-first; taking the first ``k`` clusters
    them in the morning. For ``k >= 2`` this keeps the first and last items and
    spaces the rest, so suggestions span the whole day (morning→afternoon).
    For ``k == 1`` it returns the earliest (soonest) item.
    """
    if k <= 0:
        return []
    if len(items) <= k:
        return items
    if k == 1:
        return [items[0]]
    step = (len(items) - 1) / (k - 1)
    return [items[round(i * step)] for i in range(k)]


def _emit(ctx: typer.Context, events: list[Event]) -> None:
    store_index("cal", [{"index": i, "id": e.id} for i, e in enumerate(events, start=1)])
    if ctx.obj and ctx.obj.get("json"):
        typer.echo(dump_event_list(events))
    else:
        render_event_list(Console(), events)


@cal_app.command("today")
def today_cmd(ctx: typer.Context) -> None:
    """Today's events."""
    client = GraphClient()
    events = list_events(
        client,
        start=_local_midnight(0),
        end=_local_midnight(1),
        calendar_name=None,
    )
    _emit(ctx, events)


@cal_app.command("tomorrow")
def tomorrow_cmd(ctx: typer.Context) -> None:
    """Tomorrow's events."""
    client = GraphClient()
    events = list_events(
        client, start=_local_midnight(1), end=_local_midnight(2), calendar_name=None
    )
    _emit(ctx, events)


@cal_app.command("week")
def week_cmd(
    ctx: typer.Context,
    next_week: bool = typer.Option(False, "--next", help="Next week instead of this week."),
) -> None:
    """7-day window of events."""
    offset = 7 if next_week else 0
    client = GraphClient()
    events = list_events(
        client,
        start=_local_midnight(offset),
        end=_local_midnight(offset + 7),
        calendar_name=None,
    )
    _emit(ctx, events)


@cal_app.command("list")
def list_cmd(
    ctx: typer.Context,
    start: str = typer.Option(..., "--start", help='ISO or human, e.g. "monday"'),
    end: str = typer.Option(..., "--end"),
    calendar: str | None = typer.Option(None, "--calendar"),
) -> None:
    """List events over an arbitrary range."""
    start_dt = parse_human(start)
    end_dt = parse_human(end)
    client = GraphClient()
    events = list_events(client, start=start_dt, end=end_dt, calendar_name=calendar)
    _emit(ctx, events)


def _resolve_cal_or_exit(ref: str) -> str:
    try:
        return resolve_index("cal", ref)
    except NotFound as exc:
        Console(stderr=True).print(str(exc))
        raise typer.Exit(code=64) from exc


@cal_app.command("show")
def show_cmd(
    ctx: typer.Context,
    ref: str = typer.Argument(...),
    attendees: bool = typer.Option(False, "--attendees"),
    ics: bool = typer.Option(False, "--ics", help="Output RFC 5545 ICS instead of rendered view."),
) -> None:
    """Show details of a single event."""
    event_id = _resolve_cal_or_exit(ref)
    client = GraphClient()
    ev = get_event(client, event_id=event_id)
    if ics:
        typer.echo(event_to_ics(ev))
        return
    if ctx.obj and ctx.obj.get("json"):
        typer.echo(dump_event_detail(ev))
    else:
        render_event_detail(Console(), ev, show_attendees=attendees)


def _parse_duration(text: str) -> timedelta:
    """Parse '30m', '1h', '90m', '1h30m' into timedelta."""
    pattern = re.compile(r"(\d+)\s*([hm])")
    total = timedelta()
    for match in pattern.finditer(text):
        n = int(match.group(1))
        unit = match.group(2)
        if unit == "h":
            total += timedelta(hours=n)
        else:
            total += timedelta(minutes=n)
    if total.total_seconds() == 0:
        raise typer.BadParameter(f"Cannot parse duration: {text!r}")
    return total


@cal_app.command("create")
def create_cmd(
    title: str = typer.Option(..., "--title"),
    start: str = typer.Option(..., "--start", help='ISO or human, e.g. "tomorrow 3pm"'),
    end: str | None = typer.Option(None, "--end"),
    duration: str | None = typer.Option(None, "--duration", help='"30m", "1h30m"'),
    invitees: list[str] = typer.Option([], "--invitees", help="Comma-separated or repeated."),  # noqa: B008
    location: str = typer.Option("", "--location"),
    body: str | None = typer.Option(None, "--body"),
    online: bool = typer.Option(False, "--online", help="Create as Teams meeting."),
    all_day: bool = typer.Option(False, "--all-day"),
) -> None:
    """Create a calendar event (sends invitations if --invitees provided)."""
    start_dt = parse_human(start)
    if end:
        end_dt = parse_human(end)
    elif duration:
        end_dt = start_dt + _parse_duration(duration)
    else:
        raise typer.BadParameter("Provide either --end or --duration.")
    if body is not None:
        body_text = body
    else:
        body_text = click.edit("\n# Event body. Save to use, empty to skip.\n") or ""
    body_text = body_text.split("# Event body")[0].strip()
    flat: list[str] = []
    for entry in invitees:
        flat.extend(p.strip() for p in entry.split(",") if p.strip())
    spec = CreateEvent(
        title=title,
        start=start_dt,
        end=end_dt,
        invitees=flat,
        location=location,
        body=body_text,
        is_online_meeting=online,
        is_all_day=all_day,
    )
    client = GraphClient()
    eid = create_event(client, spec)
    typer.echo(f"Created event {eid}")


@cal_app.command("respond")
def respond_cmd(
    ref: str = typer.Argument(...),
    accept: bool = typer.Option(False, "--accept"),
    decline: bool = typer.Option(False, "--decline"),
    tentative: bool = typer.Option(False, "--tentative"),
    comment: str = typer.Option("", "--comment"),
) -> None:
    """Respond to a meeting invitation."""
    chosen = [
        v for v, on in [("accept", accept), ("decline", decline), ("tentative", tentative)] if on
    ]
    if len(chosen) != 1:
        Console(stderr=True).print("Specify exactly one of --accept | --decline | --tentative.")
        raise typer.Exit(code=2)
    event_id = _resolve_cal_or_exit(ref)
    client = GraphClient()
    respond_event(client, event_id=event_id, verb=chosen[0], comment=comment)
    typer.echo(f"Responded: {chosen[0]}.")


@cal_app.command("cancel")
def cancel_cmd(
    ref: str = typer.Argument(...),
    comment: str = typer.Option("", "--comment"),
) -> None:
    """Cancel an event you organized."""
    event_id = _resolve_cal_or_exit(ref)
    client = GraphClient()
    cancel_event(client, event_id=event_id, comment=comment)
    typer.echo("Cancelled.")


@cal_app.command("find-time")
def find_time_cmd(
    ctx: typer.Context,
    with_: list[str] = typer.Option(  # noqa: B008
        [], "--with", help="Attendees (repeatable or comma-separated)."
    ),
    duration: str = typer.Option("30m", "--duration"),
    window: str = typer.Option(
        "this week",
        "--window",
        help='"this week", "next week", "today", or a date',
    ),
    per_day: int = typer.Option(
        0,
        "--per-day",
        help="Max suggestions per day (0 = all free slots). When set, slots are "
        "sampled evenly across the day.",
    ),
) -> None:
    """Suggest meeting times via Graph findMeetingTimes, across the window's days.

    Graph returns slots earliest-first, so a single wide query clusters every
    suggestion on the first free day. We probe each business day separately and
    show every free slot by default; pass --per-day N to trim each day to N
    slots sampled evenly across it (morning→afternoon).
    """
    flat: list[str] = []
    for entry in with_:
        flat.extend(p.strip() for p in entry.split(",") if p.strip())
    if not flat:
        Console(stderr=True).print("--with is required.")
        raise typer.Exit(code=2)
    if per_day < 0:
        Console(stderr=True).print("--per-day must be >= 0 (0 = all free slots).")
        raise typer.Exit(code=2)
    dur = _parse_duration(duration)
    now = datetime.now().astimezone()
    client = GraphClient()
    suggestions: list[FindTimeSuggestion] = []
    for day in _business_days(window, today=now.date()):
        win = _day_window(day, now=now)
        if win is None:
            continue
        day_slots = find_meeting_times(
            client,
            attendees=flat,
            duration=dur,
            window_start=win[0],
            window_end=win[1],
            max_candidates=_DAY_FETCH_CANDIDATES,
        )
        suggestions.extend(_spread(day_slots, per_day) if per_day > 0 else day_slots)
    if ctx.obj and ctx.obj.get("json"):
        payload = {"suggestions": [s.to_json_shape() for s in suggestions]}
        typer.echo(json.dumps(payload, indent=2, default=str))
    else:
        console = Console()
        if not suggestions:
            console.print("[dim]No suggestions returned.[/dim]")
            return
        for i, s in enumerate(suggestions, start=1):
            console.print(
                f"[{i}] {s.start.astimezone():%a %b %d %H:%M} → {s.end.astimezone():%H:%M} "
                f"(confidence: {s.confidence:.0f})"
            )
