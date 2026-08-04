"""Scrub secrets from log lines so verbose output is safe to paste."""

import re

_PATTERNS: list[tuple[re.Pattern[str], str]] = [
    (
        re.compile(r"(Authorization:\s*Bearer\s+)[A-Za-z0-9_\-.]+", re.IGNORECASE),
        r"\1***REDACTED***",
    ),
    (
        re.compile(r'("access_token"\s*:\s*")[^"]+(")'),
        r"\1***REDACTED***\2",
    ),
    (
        re.compile(r'("refresh_token"\s*:\s*")[^"]+(")'),
        r"\1***REDACTED***\2",
    ),
    (
        re.compile(r"(refresh_token=)[^&\s]+"),
        r"\1***REDACTED***",
    ),
    (
        re.compile(r"\b1\.A[A-Za-z0-9_-]{20,}"),
        "***REDACTED***",
    ),
    (
        re.compile(r"eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+"),
        "***REDACTED***",
    ),
]


def redact_secrets(text: str) -> str:
    for pattern, replacement in _PATTERNS:
        text = pattern.sub(replacement, text)
    return text
