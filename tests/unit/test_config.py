from pathlib import Path

from outlook_cli.config import cache_home, config_home, credentials_path


def test_config_home_uses_env_override(tmp_config_home: Path) -> None:
    assert config_home() == tmp_config_home


def test_cache_home_uses_env_override(tmp_cache_home: Path) -> None:
    assert cache_home() == tmp_cache_home


def test_credentials_path_lives_under_config_home(tmp_config_home: Path) -> None:
    assert credentials_path() == tmp_config_home / "credentials.json"
