"""Human and ISO date parsing for CLI flags."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

import dateparser  # type: ignore[import-untyped]

_PARSE_SETTINGS = {
    "RETURN_AS_TIMEZONE_AWARE": True,
    "PREFER_DATES_FROM": "future",
}


def parse_human(text: str) -> datetime:
    """Parse '2026-05-22T...' or 'tomorrow 3pm' or 'next monday' to a tz-aware datetime."""
    dt = dateparser.parse(text, settings=_PARSE_SETTINGS)
    if dt is None:
        raise ValueError(f"Could not parse date: '{text}'")
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=UTC)
    result: datetime = dt
    return result


def parse_since(text: str) -> datetime:
    """Parse '--since' values: '2d', '12h', 'yesterday', 'last week', or ISO."""
    text = text.strip().lower()
    now = datetime.now(UTC)
    if text.endswith("d") and text[:-1].isdigit():
        return now - timedelta(days=int(text[:-1]))
    if text.endswith("h") and text[:-1].isdigit():
        return now - timedelta(hours=int(text[:-1]))
    if text.endswith("m") and text[:-1].isdigit():
        return now - timedelta(minutes=int(text[:-1]))
    return parse_human(text)
