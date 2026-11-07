"""Inbox-zero helpers: bulk filing and the Waiting-for-reply check."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.client.mail import _conversation_items, _topic_items
from outlook_mcp.constants import OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import to_iso, truncate
from outlook_mcp.utils.quoting import split_quoted


def move_mails(
    outlook: Any,
    namespace: Any,
    *,
    entry_ids: list[str],
    target_folder: str,
    mark_read: bool | None = None,
    add_categories: list[str] | None = None,
    flag: bool | None = None,
) -> dict[str, Any]:
    """Move many messages to one folder, optionally marking/tagging them first."""
    if not entry_ids:
        raise OutlookError("Pass the entry_ids to move.")
    target = resolve_folder(namespace, target_folder)
    if add_categories:
        from outlook_mcp.client.categories import ensure_categories

        ensure_categories(namespace, add_categories, None)
    moved, errors = [], []
    for eid in entry_ids:
        try:
            item = get_item_by_id(namespace, eid)
            subject = _safe_get(item, "Subject", "")
            changed = False
            if mark_read is not None:
                item.UnRead = not mark_read
                changed = True
            if add_categories:
                from outlook_mcp.client.categories import _split

                cur = _split(_safe_get(item, "Categories", ""))
                item.Categories = ", ".join(cur + [c for c in add_categories if c not in cur])
                changed = True
            if flag is not None:
                item.FlagStatus = 2 if flag else 0
                changed = True
            if changed:
                item.Save()
            new = item.Move(target)
            moved.append({"entry_id": eid, "new_entry_id": _safe_get(new, "EntryID"), "subject": subject})
        except Exception as exc:  # noqa: BLE001 - report per message
            errors.append({"entry_id": eid, "error": str(exc)})
    out = {"status": "moved", "folder": _safe_get(target, "Name"), "count": len(moved), "items": moved}
    if errors:
        out["errors"] = errors
    return out


# ----------------------------------------------------------------- waiting


def _naive(value: Any) -> dt.datetime | None:
    if value is None or not hasattr(value, "year"):
        return None
    return dt.datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)


def _me_keys(namespace: Any) -> set[str]:
    me = _safe_get(namespace, "CurrentUser")
    keys = {str(_safe_get(me, "Name", "") or "").lower()}
    try:
        keys.add(str(me.AddressEntry.GetExchangeUser().PrimarySmtpAddress).lower())
    except Exception:
        keys.add(str(_safe_get(me, "Address", "") or "").lower())
    return {k for k in keys if k}


def _from_me(item: Any, me: set[str]) -> bool:
    sender = {
        str(_safe_get(item, p, "") or "").lower()
        for p in ("SenderName", "SenderEmailAddress", "_sender_smtp")
    }
    return bool(sender & me)


def _waiting_folder(namespace: Any, folder: str | None) -> tuple[Any, str]:
    if folder:
        return resolve_folder(namespace, folder), "given"
    from outlook_mcp.client.quicksteps import quick_step_target

    hit = quick_step_target(namespace, "wait")
    if hit is not None:
        return hit[1], "quick_step"
    candidates = []
    inbox = resolve_folder(namespace, "inbox")
    for parent in (inbox, _safe_get(inbox, "Parent")):
        for sub in list(_safe_get(parent, "Folders", []) or []):
            if "wait" in str(_safe_get(sub, "Name", "")).lower():
                candidates.append(sub)
    if len(candidates) == 1:
        return candidates[0], "folder_name"
    raise OutlookError(
        "Couldn't tell which folder is your Waiting folder. Pass folder=... "
        "(e.g. 'Inbox/2-Waiting')."
    )


def check_waiting(
    outlook: Any,
    namespace: Any,
    *,
    folder: str | None = None,
    move_answered_to: str | None = None,
    limit: int = 200,
) -> dict[str, Any]:
    """Which messages in the Waiting folder got a reply since they were filed."""
    wait_folder, how = _waiting_folder(namespace, folder)
    me = _me_keys(namespace)
    now = dt.datetime.now()
    answered, waiting = [], []
    n = 0
    for item in list(wait_folder.Items):
        if _safe_get(item, "Class") not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        n += 1
        if n > limit:
            break
        eid = str(_safe_get(item, "EntryID", ""))
        since = _naive(_safe_get(item, "ReceivedTime")) or _naive(_safe_get(item, "SentOn"))
        thread = _conversation_items(item)
        if thread is None:
            thread = _topic_items(namespace, item)
        replies = []
        for other in thread:
            if str(_safe_get(other, "EntryID", "")) == eid:
                continue
            if _safe_get(other, "Class") not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
                continue
            if not _safe_get(other, "Sent", True) or _from_me(other, me):
                continue
            when = _naive(_safe_get(other, "ReceivedTime")) or _naive(_safe_get(other, "SentOn"))
            if since and when and when > since:
                body, _ = split_quoted(str(_safe_get(other, "Body", "") or ""))
                replies.append({
                    "entry_id": _safe_get(other, "EntryID"),
                    "from": _safe_get(other, "SenderName"),
                    "when": to_iso(_safe_get(other, "ReceivedTime")),
                    "preview": truncate(body, 200),
                    "_when": when,
                })
        row = {
            "entry_id": eid,
            "subject": _safe_get(item, "Subject", ""),
            "to": _safe_get(item, "To", ""),
            "waiting_since": to_iso(_safe_get(item, "ReceivedTime") or _safe_get(item, "SentOn")),
            "days_waiting": (now - since).days if since else None,
        }
        if replies:
            replies.sort(key=lambda r: r["_when"])
            row["replies"] = [{k: v for k, v in r.items() if k != "_when"} for r in replies]
            answered.append(row)
        else:
            waiting.append(row)
    waiting.sort(key=lambda r: -(r["days_waiting"] or 0))
    out: dict[str, Any] = {
        "folder": _safe_get(wait_folder, "Name"),
        "folder_found_by": how,
        "answered_count": len(answered),
        "still_waiting_count": len(waiting),
        "answered": answered,
        "still_waiting": waiting,
    }
    if move_answered_to and answered:
        out["moved"] = move_mails(
            outlook, namespace, entry_ids=[a["entry_id"] for a in answered], target_folder=move_answered_to
        )
    return out
