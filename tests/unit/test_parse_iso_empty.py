"""Verify _parse_iso raises on empty input (no silent now() fallback)."""

import pytest

from outlook_cli.graph.models import _parse_iso


def test_parse_iso_raises_on_empty_string() -> None:
    with pytest.raises(ValueError, match="empty"):
        _parse_iso("")


def test_parse_iso_accepts_iso_with_z() -> None:
    dt = _parse_iso("2026-05-22T10:00:00Z")
    assert dt.year == 2026 and dt.hour == 10


def test_parse_iso_accepts_iso_with_microseconds() -> None:
    dt = _parse_iso("2026-05-22T10:00:00.123456Z")
    assert dt.microsecond == 123456
