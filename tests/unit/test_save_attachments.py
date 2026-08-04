"""Path-traversal rejection tests for commands/mail._save_attachments."""

from pathlib import Path
from typing import Any
from unittest.mock import MagicMock

from outlook_cli.commands.mail import _save_attachments


def _att(name: str, content_b64: str = "aGVsbG8=") -> dict[str, Any]:
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentBytes": content_b64,
    }


def test_save_attachments_strips_directory_components(tmp_path: Path) -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": [_att("../escaped.txt")]}
    _save_attachments(client, "MSG-1", str(tmp_path))
    # File should be saved as "escaped.txt" (basename only) inside tmp_path
    assert (tmp_path / "escaped.txt").exists()
    # No file should have escaped the directory
    parent = tmp_path.parent
    assert not (parent / "escaped.txt").exists()


def test_save_attachments_rejects_absolute_path_escape(tmp_path: Path) -> None:
    # An absolute-looking name's basename is just the leaf, which is safe.
    # But if Path.name normalization ever changes, we want a regression test.
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": [_att("/etc/passwd")]}
    _save_attachments(client, "MSG-1", str(tmp_path))
    assert (tmp_path / "passwd").exists()
    assert not Path("/etc/passwd-outlook-cli-test").exists()  # negative check


def test_save_attachments_skips_empty_filename(tmp_path: Path) -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {"value": [_att(""), _att("real.txt")]}
    _save_attachments(client, "MSG-1", str(tmp_path))
    assert (tmp_path / "real.txt").exists()
    # Empty-name attachment should have been silently skipped
    assert len(list(tmp_path.iterdir())) == 1


def test_save_attachments_skips_non_file_attachments(tmp_path: Path) -> None:
    client = MagicMock()
    client.get.return_value.json.return_value = {
        "value": [
            {"@odata.type": "#microsoft.graph.itemAttachment", "name": "ref.eml"},
            _att("real.txt"),
        ]
    }
    _save_attachments(client, "MSG-1", str(tmp_path))
    assert (tmp_path / "real.txt").exists()
    assert not (tmp_path / "ref.eml").exists()
