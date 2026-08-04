"""Filesystem paths and environment overrides for outlook-cli."""

import os
import ssl
from pathlib import Path


def _expand(env_var: str, fallback: Path) -> Path:
    override = os.environ.get(env_var)
    return Path(override) if override else fallback


def config_home() -> Path:
    return _expand("OUTLOOK_CLI_CONFIG_HOME", Path.home() / ".config" / "outlook-cli")


def cache_home() -> Path:
    return _expand("OUTLOOK_CLI_CACHE_HOME", Path.home() / ".cache" / "outlook-cli")


def credentials_path() -> Path:
    return config_home() / "credentials.json"


def access_tokens_path() -> Path:
    return cache_home() / "access_tokens.json"


def folders_cache_path() -> Path:
    return cache_home() / "folders.json"


def mail_index_path() -> Path:
    return cache_home() / "last_mail_listing.json"


def cal_index_path() -> Path:
    return cache_home() / "last_cal_listing.json"


def ensure_dirs() -> None:
    config_home().mkdir(parents=True, exist_ok=True)
    cache_home().mkdir(parents=True, exist_ok=True)


def http_verify() -> bool | ssl.SSLContext | str:
    """Resolve the TLS verification setting for outgoing HTTPS requests.

    Resolution order:
      1. ``OUTLOOK_CLI_INSECURE=1`` -> disable verification entirely.
      2. ``OUTLOOK_CLI_CA_BUNDLE`` / ``REQUESTS_CA_BUNDLE`` / ``SSL_CERT_FILE``
         -> use that file/dir as the trust store (handles corporate MITM proxies).
      3. ``truststore`` -> use the OS trust store (picks up enterprise roots
         installed in Windows/macOS cert stores).
      4. ``certifi`` -> fall back to certifi's bundled CA list.
      5. Default ``True`` -> use whatever ssl considers the system store.
    """
    if os.environ.get("OUTLOOK_CLI_INSECURE", "").lower() in ("1", "true", "yes"):
        return False
    for var in ("OUTLOOK_CLI_CA_BUNDLE", "REQUESTS_CA_BUNDLE", "SSL_CERT_FILE"):
        bundle = os.environ.get(var)
        if bundle:
            return bundle
    try:
        import truststore

        return truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
    except ImportError:
        pass
    try:
        import certifi

        return certifi.where()
    except ImportError:
        return True
