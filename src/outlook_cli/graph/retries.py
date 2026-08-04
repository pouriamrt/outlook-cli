"""Retry policy for Graph HTTP calls: 429 honors Retry-After, 5xx exp backoff, 401 force-refresh."""

from __future__ import annotations

import logging
import random
import time
from collections.abc import Callable
from typing import Any

import httpx

logger = logging.getLogger(__name__)

DEFAULT_MAX_ATTEMPTS = 3
DEFAULT_BACKOFF_BASE = 1.0
DEFAULT_BACKOFF_MAX = 30.0


def _backoff(attempt: int, base: float, cap: float) -> float:
    raw: float = min(base * (2 ** (attempt - 1)), cap)
    jitter: float = raw * random.uniform(-0.2, 0.2)
    return float(max(0.0, raw + jitter))


def with_retries(
    fn: Callable[..., httpx.Response],
    *,
    refresh_token_fn: Callable[[], None],
    max_attempts: int = DEFAULT_MAX_ATTEMPTS,
    backoff_base: float = DEFAULT_BACKOFF_BASE,
    backoff_max: float = DEFAULT_BACKOFF_MAX,
) -> Callable[..., httpx.Response]:
    """Wrap a Graph call with retry semantics."""

    def wrapped(*args: Any, **kwargs: Any) -> httpx.Response:
        attempt = 0
        used_force_refresh = False
        while True:
            attempt += 1
            try:
                resp = fn(*args, **kwargs)
            except httpx.TransportError as exc:
                if attempt >= max_attempts:
                    raise
                delay = _backoff(attempt, backoff_base, backoff_max)
                logger.info("Network error %s; retry %d after %.1fs", exc, attempt, delay)
                time.sleep(delay)
                continue

            if resp.status_code == 401 and not used_force_refresh:
                used_force_refresh = True
                logger.info("401 from Graph; forcing token refresh and retrying once")
                refresh_token_fn()
                continue

            if resp.status_code == 429 and attempt < max_attempts:
                retry_after = min(float(resp.headers.get("Retry-After", "5")), backoff_max)
                logger.info("429 throttled; retry %d after %.1fs", attempt, retry_after)
                time.sleep(retry_after)
                continue

            if 500 <= resp.status_code < 600 and attempt < max_attempts:
                delay = _backoff(attempt, backoff_base, backoff_max)
                logger.info("%d from Graph; retry %d after %.1fs", resp.status_code, attempt, delay)
                time.sleep(delay)
                continue

            return resp

    return wrapped
