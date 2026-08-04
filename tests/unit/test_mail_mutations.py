from unittest.mock import MagicMock

from outlook_cli.graph.mail import (
    delete_message,
    flag_message,
    mark_message,
    move_message,
    unflag_message,
)


def test_move_posts_to_move_with_destination() -> None:
    client = MagicMock()
    client.post.return_value.json.return_value = {"id": "NEW-ID"}
    new_id = move_message(client, message_id="M1", destination_folder_id="archive")
    args, kwargs = client.post.call_args
    assert "/me/messages/M1/move" in args[0]
    assert kwargs["json_body"] == {"destinationId": "archive"}
    assert new_id == "NEW-ID"


def test_delete_uses_DELETE_for_soft() -> None:  # noqa: N802
    client = MagicMock()
    delete_message(client, message_id="M1", purge=False)
    args, kwargs = client.delete.call_args
    assert "/me/messages/M1" in args[0]


def test_delete_purge_first_moves_to_deleted_then_deletes() -> None:
    client = MagicMock()
    delete_message(client, message_id="M1", purge=True)
    assert client.post.called
    assert client.delete.called


def test_flag_patches_flag_status() -> None:
    client = MagicMock()
    flag_message(client, message_id="M1", due=None)
    args, kwargs = client.patch.call_args
    assert "/me/messages/M1" in args[0]
    assert kwargs["json_body"]["flag"]["flagStatus"] == "flagged"


def test_unflag_patches_to_notFlagged() -> None:  # noqa: N802
    client = MagicMock()
    unflag_message(client, message_id="M1")
    assert client.patch.call_args.kwargs["json_body"]["flag"]["flagStatus"] == "notFlagged"


def test_mark_read_patches_isRead_true() -> None:  # noqa: N802
    client = MagicMock()
    mark_message(client, message_id="M1", read=True)
    assert client.patch.call_args.kwargs["json_body"] == {"isRead": True}


def test_mark_unread_patches_isRead_false() -> None:  # noqa: N802
    client = MagicMock()
    mark_message(client, message_id="M1", read=False)
    assert client.patch.call_args.kwargs["json_body"] == {"isRead": False}
