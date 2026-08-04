"""Shared pytest fixtures for outlook-cli."""

from pathlib import Path

import pytest


@pytest.fixture
def tmp_config_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Redirect ~/.config/outlook-cli/ to a tmp dir for the test."""
    cfg = tmp_path / "config" / "outlook-cli"
    cfg.mkdir(parents=True)
    monkeypatch.setenv("OUTLOOK_CLI_CONFIG_HOME", str(cfg))
    return cfg


@pytest.fixture
def tmp_cache_home(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """Redirect ~/.cache/outlook-cli/ to a tmp dir for the test."""
    cache = tmp_path / "cache" / "outlook-cli"
    cache.mkdir(parents=True)
    monkeypatch.setenv("OUTLOOK_CLI_CACHE_HOME", str(cache))
    return cache
