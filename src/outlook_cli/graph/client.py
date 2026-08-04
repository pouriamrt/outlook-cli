"""HTTP client for Microsoft Graph / Outlook REST with bearer injection and retry orchestration."""

from __future__ import annotations

import logging
import os
from collections.abc import Callable
from typing import Any

import httpx

from outlook_cli.auth.token_refresh import GRAPH_SCOPE, get_token
from outlook_cli.config import http_verify

logger = logging.getLogger(__name__)

GRAPH_BASE_URL = "https://outlook.office.com/api/v2.0"


def _default_base_url() -> str:
    """Resolve the API base URL, honouring ``OUTLOOK_CLI_API_BASE``.

    Set ``OUTLOOK_CLI_API_BASE=https://outlook.office.com/api/v2.0`` when the
    captured refresh token is only consented for the Outlook audience.
    """
    return os.environ.get("OUTLOOK_CLI_API_BASE", GRAPH_BASE_URL)


def _default_scope() -> str:
    return os.environ.get("OUTLOOK_CLI_API_SCOPE", GRAPH_SCOPE)


def _is_outlook_rest_base(base_url: str) -> bool:
    """Outlook REST v2.0 (outlook.office.com) wants PascalCase JSON; Graph wants camelCase."""
    return "outlook.office.com" in base_url


def _denormalize_keys(data: Any) -> Any:
    """PascalCase every dict key for outgoing Outlook REST payloads.

    Mirror of ``_normalize_keys`` for responses. Keys starting with ``@``
    (OData metadata like ``@odata.type``, ``@odata.bind``) are preserved
    literally.
    """
    if isinstance(data, dict):
        return {
            (k if k.startswith("@") else k[0].upper() + k[1:] if k else k): (_denormalize_keys(v))
            for k, v in data.items()
        }
    if isinstance(data, list):
        return [_denormalize_keys(item) for item in data]
    return data


class GraphError(Exception):
    def __init__(self, status_code: int, body: Any) -> None:
        super().__init__(f"Graph error {status_code}: {body}")
        self.status_code = status_code
        self.body = body


class GraphClient:
    def __init__(
        self,
        *,
        token_provider: Callable[[str], str] = get_token,
        scope: str | None = None,
        base_url: str | None = None,
        timeout: float = 30.0,
    ) -> None:
        self._token_provider = token_provider
        self._scope = scope if scope is not None else _default_scope()
        resolved_base = base_url if base_url is not None else _default_base_url()
        self._is_outlook_rest = _is_outlook_rest_base(resolved_base)
        self._client = httpx.Client(base_url=resolved_base, timeout=timeout, verify=http_verify())

    def _prep_body(self, json_body: Any) -> Any:
        """PascalCase keys when targeting Outlook REST; leave Graph payloads as-is."""
        if json_body is None or not self._is_outlook_rest:
            return json_body
        return _denormalize_keys(json_body)

    def _headers(self) -> dict[str, str]:
        token = self._token_provider(self._scope)
        return {"Authorization": f"Bearer {token}"}

    def _wrap(self, fn: Callable[..., httpx.Response]) -> Callable[..., httpx.Response]:
        from outlook_cli.graph.retries import with_retries

        def _force_refresh() -> None:
            from outlook_cli.config import access_tokens_path

            access_tokens_path().unlink(missing_ok=True)

        return with_retries(fn, refresh_token_fn=_force_refresh)

    def get(self, path: str, *, params: dict[str, Any] | None = None) -> httpx.Response:
        def _call() -> httpx.Response:
            return self._client.get(path, params=params, headers=self._headers())

        return _check(self._wrap(_call)())

    def post(
        self,
        path: str,
        *,
        json_body: Any = None,
        params: dict[str, Any] | None = None,
    ) -> httpx.Response:
        body = self._prep_body(json_body)

        def _call() -> httpx.Response:
            return self._client.post(path, json=body, params=params, headers=self._headers())

        return _check(self._wrap(_call)())

    def patch(self, path: str, *, json_body: Any = None) -> httpx.Response:
        body = self._prep_body(json_body)

        def _call() -> httpx.Response:
            return self._client.patch(path, json=body, headers=self._headers())

        return _check(self._wrap(_call)())

    def delete(self, path: str) -> httpx.Response:
        def _call() -> httpx.Response:
            return self._client.delete(path, headers=self._headers())

        return _check(self._wrap(_call)())

    def close(self) -> None:
        self._client.close()

    def __enter__(self) -> GraphClient:
        return self

    def __exit__(self, exc_type: object, exc_val: object, exc_tb: object) -> None:
        self.close()


def _normalize_keys(data: Any) -> Any:
    if isinstance(data, dict):
        return {(k[0].lower() + k[1:] if k else k): _normalize_keys(v) for k, v in data.items()}
    elif isinstance(data, list):
        return [_normalize_keys(item) for item in data]
    return data


def _check(resp: httpx.Response) -> httpx.Response:
    if resp.status_code >= 400:
        try:
            body: Any = resp.json()
        except Exception:
            body = resp.text
        raise GraphError(resp.status_code, body)

    orig_json = resp.json

    def normalized_json(*args: Any, **kwargs: Any) -> Any:
        return _normalize_keys(orig_json(*args, **kwargs))

    resp.json = normalized_json  # type: ignore[method-assign]

    return resp
