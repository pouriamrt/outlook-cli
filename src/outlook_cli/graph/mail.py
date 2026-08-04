"""Mail primitives on Microsoft Graph. Typed args → typed results."""

from __future__ import annotations

import base64 as _base64
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx

from outlook_cli.config import http_verify
from outlook_cli.graph.client import GraphClient, GraphError
from outlook_cli.graph.models import Message

_MESSAGE_SELECT = (
    "id,subject,from,toRecipients,ccRecipients,receivedDateTime,"
    "isRead,flag,hasAttachments,importance,bodyPreview,conversationId"
)


@dataclass
class MailListFilters:
    unread: bool = False
    flagged: bool = False
    from_addr: str | None = None
    subject: str | None = None
    since: datetime | None = None
    top: int = 25
    skip: int = 0


@dataclass
class MailListResult:
    items: list[Message] = field(default_factory=list)
    next_link: str | None = None


def build_list_params(f: MailListFilters) -> dict[str, Any]:
    clauses: list[str] = []
    if f.unread:
        clauses.append("isRead eq false")
    if f.flagged:
        clauses.append("flag/flagStatus eq 'flagged'")
    if f.from_addr:
        escaped_addr = f.from_addr.replace("'", "''")
        clauses.append(f"from/emailAddress/address eq '{escaped_addr}'")
    if f.subject:
        escaped = f.subject.replace("'", "''")
        clauses.append(f"contains(subject, '{escaped}')")
    if f.since:
        iso = f.since.strftime("%Y-%m-%dT%H:%M:%SZ")
        clauses.append(f"receivedDateTime ge {iso}")
    params: dict[str, Any] = {
        "$top": f.top,
        "$orderby": "receivedDateTime desc",
        "$select": _MESSAGE_SELECT,
    }
    if clauses:
        params["$filter"] = " and ".join(clauses)
    if f.skip:
        params["$skip"] = f.skip
    return params


def list_messages(
    client: GraphClient,
    *,
    folder_id: str,
    filters: MailListFilters,
) -> MailListResult:
    params = build_list_params(filters)
    resp = client.get(f"/me/mailFolders/{folder_id}/messages", params=params)
    payload: dict[str, Any] = resp.json()
    items = [Message.from_graph(m) for m in payload.get("value", [])]
    next_link: str | None = payload.get("@odata.nextLink")
    return MailListResult(items=items, next_link=next_link)


def read_message(client: GraphClient, *, message_id: str) -> Message:
    resp = client.get(
        f"/me/messages/{message_id}",
        params={"$select": _MESSAGE_SELECT + ",body"},
    )
    return Message.from_graph(resp.json())


def read_thread(client: GraphClient, *, conversation_id: str) -> list[Message]:
    """Return all messages in a conversation, oldest first."""
    escaped = conversation_id.replace("'", "''")
    resp = client.get(
        "/me/messages",
        params={
            "$filter": f"conversationId eq '{escaped}'",
            "$orderby": "receivedDateTime asc",
            "$top": 100,
            "$select": _MESSAGE_SELECT + ",body",
        },
    )
    messages = [Message.from_graph(m) for m in resp.json().get("value", [])]
    messages.sort(key=lambda m: m.received_at)
    return messages


_INLINE_ATTACH_LIMIT = 3 * 1024 * 1024  # 3 MB


@dataclass
class SendDraft:
    to: list[str]
    subject: str
    body: str
    body_html: bool = True
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    attachments: list[Path] = field(default_factory=list)
    importance: str = "normal"
    save_as_draft: bool = False
    reply_to_id: str | None = None


def _addr_list(addrs: list[str]) -> list[dict[str, Any]]:
    return [{"emailAddress": {"address": a}} for a in addrs]


def _build_message(draft: SendDraft) -> dict[str, Any]:
    msg: dict[str, Any] = {
        "subject": draft.subject,
        "body": {
            "contentType": "HTML" if draft.body_html else "Text",
            "content": draft.body,
        },
        "toRecipients": _addr_list(draft.to),
        "ccRecipients": _addr_list(draft.cc),
        "bccRecipients": _addr_list(draft.bcc),
        "importance": draft.importance,
    }
    small = [a for a in draft.attachments if a.stat().st_size <= _INLINE_ATTACH_LIMIT]
    if small:
        msg["attachments"] = [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": a.name,
                "contentBytes": _base64.b64encode(a.read_bytes()).decode("ascii"),
            }
            for a in small
        ]
    return msg


def _upload_large_attachments(client: GraphClient, message_id: str, files: list[Path]) -> None:
    for f in files:
        size = f.stat().st_size
        sess = client.post(
            f"/me/messages/{message_id}/attachments/createUploadSession",
            json_body={
                "AttachmentItem": {
                    "attachmentType": "file",
                    "name": f.name,
                    "size": size,
                }
            },
        )
        upload_url: str = sess.json()["uploadUrl"]
        chunk = 5 * 1024 * 1024  # 5 MB chunks
        with f.open("rb") as fh, httpx.Client(timeout=120.0, verify=http_verify()) as raw:
            offset = 0
            while offset < size:
                data = fh.read(chunk)
                end = offset + len(data) - 1
                put_resp = raw.put(
                    upload_url,
                    content=data,
                    headers={
                        "Content-Length": str(len(data)),
                        "Content-Range": f"bytes {offset}-{end}/{size}",
                    },
                )
                if put_resp.status_code >= 400:
                    raise GraphError(put_resp.status_code, put_resp.text)
                offset = end + 1


def send_mail(client: GraphClient, draft: SendDraft) -> str | None:
    """Send a message (or save as draft). Returns message_id when save_as_draft is True."""
    large = [a for a in draft.attachments if a.stat().st_size > _INLINE_ATTACH_LIMIT]
    if draft.save_as_draft or large:
        created = client.post("/me/messages", json_body=_build_message(draft))
        message_id: str = created.json()["id"]
        if large:
            _upload_large_attachments(client, message_id, large)
        if not draft.save_as_draft:
            client.post(f"/me/messages/{message_id}/send")
        return message_id
    client.post(
        "/me/sendMail",
        json_body={"message": _build_message(draft), "saveToSentItems": True},
    )
    return None


def _patch_draft_body(client: GraphClient, draft_id: str, draft: SendDraft) -> None:
    """Update the auto-created reply/forward draft with our subject/body/recipients."""
    patch_payload: dict[str, Any] = {
        "body": {
            "contentType": "HTML" if draft.body_html else "Text",
            "content": draft.body,
        },
    }
    if draft.subject:
        patch_payload["subject"] = draft.subject
    if draft.to:
        patch_payload["toRecipients"] = _addr_list(draft.to)
    if draft.cc:
        patch_payload["ccRecipients"] = _addr_list(draft.cc)
    if draft.bcc:
        patch_payload["bccRecipients"] = _addr_list(draft.bcc)
    client.patch(f"/me/messages/{draft_id}", json_body=patch_payload)


def reply_mail(client: GraphClient, draft: SendDraft, *, reply_all: bool) -> None:
    if not draft.reply_to_id:
        raise ValueError("reply_to_id is required for reply_mail")
    endpoint = "createReplyAll" if reply_all else "createReply"
    created = client.post(f"/me/messages/{draft.reply_to_id}/{endpoint}")
    draft_id: str = created.json()["id"]
    _patch_draft_body(client, draft_id, draft)
    if draft.attachments:
        large = [a for a in draft.attachments if a.stat().st_size > _INLINE_ATTACH_LIMIT]
        small = [a for a in draft.attachments if a.stat().st_size <= _INLINE_ATTACH_LIMIT]
        for a in small:
            client.post(
                f"/me/messages/{draft_id}/attachments",
                json_body={
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": a.name,
                    "contentBytes": _base64.b64encode(a.read_bytes()).decode("ascii"),
                },
            )
        if large:
            _upload_large_attachments(client, draft_id, large)
    if not draft.save_as_draft:
        client.post(f"/me/messages/{draft_id}/send")


def forward_mail(client: GraphClient, draft: SendDraft) -> None:
    if not draft.reply_to_id:
        raise ValueError("reply_to_id is required for forward_mail")
    if not draft.to:
        raise ValueError("Forward requires at least one recipient in draft.to")
    created = client.post(f"/me/messages/{draft.reply_to_id}/createForward")
    draft_id: str = created.json()["id"]
    _patch_draft_body(client, draft_id, draft)
    if not draft.save_as_draft:
        client.post(f"/me/messages/{draft_id}/send")


def move_message(client: GraphClient, *, message_id: str, destination_folder_id: str) -> str:
    resp = client.post(
        f"/me/messages/{message_id}/move",
        json_body={"destinationId": destination_folder_id},
    )
    new_id: str = resp.json()["id"]
    return new_id


def delete_message(client: GraphClient, *, message_id: str, purge: bool) -> None:
    if purge:
        moved_id = move_message(client, message_id=message_id, destination_folder_id="deleteditems")
        client.delete(f"/me/messages/{moved_id}")
    else:
        client.delete(f"/me/messages/{message_id}")


def flag_message(client: GraphClient, *, message_id: str, due: datetime | None) -> None:
    flag: dict[str, Any] = {"flagStatus": "flagged"}
    if due is not None:
        flag["dueDateTime"] = {
            "dateTime": due.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": "UTC",
        }
    client.patch(f"/me/messages/{message_id}", json_body={"flag": flag})


def unflag_message(client: GraphClient, *, message_id: str) -> None:
    client.patch(
        f"/me/messages/{message_id}",
        json_body={"flag": {"flagStatus": "notFlagged"}},
    )


def mark_message(client: GraphClient, *, message_id: str, read: bool) -> None:
    client.patch(f"/me/messages/{message_id}", json_body={"isRead": read})


def search_messages(
    client: GraphClient,
    *,
    query: str,
    folder_id: str | None,
    top: int = 25,
) -> list[Message]:
    """KQL-style search via $search. Note: $orderby is not allowed with $search."""
    base = "/me/messages" if not folder_id else f"/me/mailFolders/{folder_id}/messages"
    resp = client.get(
        base,
        params={
            "$search": f'"{query}"',
            "$top": top,
            "$select": _MESSAGE_SELECT,
        },
    )
    return [Message.from_graph(m) for m in resp.json().get("value", [])]
