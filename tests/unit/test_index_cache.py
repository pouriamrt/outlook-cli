from pathlib import Path

import pytest

from outlook_cli.errors import NotFound
from outlook_cli.index_cache import resolve, store


def test_store_and_resolve_mail_index(tmp_cache_home: Path) -> None:
    store("mail", [{"index": 1, "id": "ID-1"}, {"index": 2, "id": "ID-2"}])
    assert resolve("mail", 1) == "ID-1"
    assert resolve("mail", 2) == "ID-2"


def test_resolve_raises_when_index_missing(tmp_cache_home: Path) -> None:
    store("mail", [{"index": 1, "id": "ID-1"}])
    with pytest.raises(NotFound) as exc:
        resolve("mail", 99)
    assert "99" in str(exc.value)


def test_resolve_passes_through_graph_id(tmp_cache_home: Path) -> None:
    full_id = "AAMkAGI2NzNkY2I5LWFiY2QtMTIzNA=="
    assert resolve("mail", full_id) == full_id


def test_mail_and_cal_caches_are_independent(tmp_cache_home: Path) -> None:
    store("mail", [{"index": 1, "id": "MAIL-1"}])
    store("cal", [{"index": 1, "id": "CAL-1"}])
    assert resolve("mail", 1) == "MAIL-1"
    assert resolve("cal", 1) == "CAL-1"


def test_resolve_raises_when_cache_missing(tmp_cache_home: Path) -> None:
    with pytest.raises(NotFound) as exc:
        resolve("mail", 1)
    assert "outlook mail list" in str(exc.value)
