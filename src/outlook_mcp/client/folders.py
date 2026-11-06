"""Folder resolution + list/create.

All public functions take ``(outlook, namespace, **kwargs)`` so they can
run on the bridge thread directly. Internal helpers take ``namespace``
since they don't need the application object.
"""

from __future__ import annotations

from typing import Any

import pythoncom

from outlook_mcp.constants import DEFAULT_FOLDER_MAP, OL_FOLDER_INBOX
from outlook_mcp.errors import OutlookError, is_disconnect_error


def _safe_get(item: Any, attr: str, default: Any = None) -> Any:
    try:
        return getattr(item, attr)
    except Exception:
        return default


def resolve_folder(namespace: Any, folder: str | None) -> Any:
    """Resolve a folder spec to a MAPIFolder COM object.

    Accepted forms:
      * None / "" -> default Inbox
      * one of the well-known names (``inbox``, ``sent``, ...)
      * a slash-delimited path (``Inbox/Projects/Quinn``), optionally
        prefixed with a store display name.
    """
    if not folder:
        return namespace.GetDefaultFolder(OL_FOLDER_INBOX)

    key = folder.strip().lower()
    if key in DEFAULT_FOLDER_MAP:
        return namespace.GetDefaultFolder(DEFAULT_FOLDER_MAP[key])

    segments = [s for s in folder.replace("\\", "/").split("/") if s]
    if not segments:
        raise OutlookError(f"Invalid folder spec: '{folder}'")

    first = segments[0].lower()
    root = None
    for store in namespace.Stores:
        if store.DisplayName.lower() == first:
            root = store.GetRootFolder()
            segments = segments[1:]
            break
    if root is None:
        root = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent

    current = root
    for seg in segments:
        seg_lower = seg.lower()
        match = None
        for sub in current.Folders:
            if sub.Name.lower() == seg_lower:
                match = sub
                break
        if match is None:
            raise OutlookError(
                f"Folder '{seg}' not found under '{current.Name}'. "
                "Use outlook_list_folders to see available paths."
            )
        current = match
    return current


def get_item_by_id(namespace: Any, entry_id: str, store_id: str | None = None) -> Any:
    try:
        if store_id:
            return namespace.GetItemFromID(entry_id, store_id)
        return namespace.GetItemFromID(entry_id)
    except pythoncom.com_error as exc:
        if is_disconnect_error(exc):
            # Outlook itself is gone, not the item — let the bridge see
            # the raw COM error so it reconnects and retries.
            raise
        raise OutlookError(
            f"Item with id '{entry_id}' not found. The id may be stale or "
            "the item may have been deleted."
        ) from exc


def list_folders(outlook: Any, namespace: Any, *, root: str | None = None, max_depth: int = 4) -> list[dict[str, Any]]:
    if root:
        start = resolve_folder(namespace, root)
    else:
        start = namespace.GetDefaultFolder(OL_FOLDER_INBOX).Parent

    out: list[dict[str, Any]] = []

    def walk(folder: Any, path: str, depth: int) -> None:
        items_obj = _safe_get(folder, "Items")
        out.append(
            {
                "name": folder.Name,
                "path": path,
                "item_count": items_obj.Count if items_obj else 0,
                "unread_count": _safe_get(folder, "UnReadItemCount", 0),
                "default_item_type": _safe_get(folder, "DefaultItemType", -1),
            }
        )
        if depth >= max_depth:
            return
        for sub in folder.Folders:
            walk(sub, f"{path}/{sub.Name}", depth + 1)

    walk(start, start.Name, 0)
    return out


def create_folder(outlook: Any, namespace: Any, *, parent: str | None, name: str) -> dict[str, Any]:
    parent_folder = resolve_folder(namespace, parent)
    new_folder = parent_folder.Folders.Add(name)
    return {
        "name": new_folder.Name,
        "path": f"{parent_folder.Name}/{new_folder.Name}",
        "entry_id": new_folder.EntryID,
    }


# --------------------------------------------------------------------------
# Mailboxes (stores) and multi-folder search scopes
# --------------------------------------------------------------------------

# OlExchangeStoreType
_STORE_TYPES = {
    0: "primary",  # olPrimaryExchangeMailbox
    1: "delegate",  # olExchangeMailbox (opened as delegate)
    2: "public_folders",  # olExchangePublicFolder
    3: "pst_or_other",  # olNotExchange (PST, IMAP, ...)
    4: "shared",  # olAdditionalExchangeMailbox
}

# Matched against the store's display name ("Online Archive - you@corp.com",
# "In-Place Archive - ...", "Online-Archiv - ..." in German, "ארכיון מקוון -"
# in Hebrew, "Archivo en línea -" in Spanish ...).
_ARCHIVE_MARKERS = ("archive", "archiv", "ארכיון")

# Folders a mail search should not descend into.
_SKIP_FOLDER_NAMES = {
    "sync issues", "conflicts", "local failures", "server failures",
    "conversation history", "rss feeds", "rss subscriptions", "outbox",
    "conversation action settings", "quick step settings", "yammer root",
    "files", "recoverable items",
}
_SKIP_WHEN_EXCLUDING_DELETED = {"deleted items", "junk email", "junk e-mail", "trash", "junk"}

MAIL_ITEM_TYPE = 0  # Folder.DefaultItemType for mail folders


def store_kind(store: Any) -> str:
    name = str(_safe_get(store, "DisplayName", "") or "").lower()
    if any(m in name for m in _ARCHIVE_MARKERS):
        return "archive"
    if _safe_get(store, "IsArchive", False):  # not in every build; cheap to try
        return "archive"
    return _STORE_TYPES.get(_safe_get(store, "ExchangeStoreType", 3), "pst_or_other")


def list_stores(namespace: Any) -> list[dict[str, Any]]:
    default_id = _safe_get(_safe_get(namespace, "DefaultStore"), "StoreID")
    out = []
    for store in namespace.Stores:
        out.append(
            {
                "name": str(_safe_get(store, "DisplayName", "")),
                "kind": store_kind(store),
                "is_default": bool(default_id and _safe_get(store, "StoreID") == default_id),
            }
        )
    return out


def list_mailboxes(outlook: Any, namespace: Any) -> dict[str, Any]:
    stores = list_stores(namespace)
    return {
        "count": len(stores),
        "mailboxes": stores,
        "note": (
            "Pass search_in='all' (every mailbox), 'archive', or 'mailbox' "
            "(your own), or name specific ones with mailboxes=[...]."
        ),
    }


def select_stores(namespace: Any, search_in: str, mailboxes: list[str] | None) -> list[Any]:
    default_id = _safe_get(_safe_get(namespace, "DefaultStore"), "StoreID")
    stores = list(namespace.Stores)
    if mailboxes:
        wanted = [m.strip().lower() for m in mailboxes if m.strip()]
        picked = [s for s in stores if any(w in str(s.DisplayName).lower() for w in wanted)]
        if not picked:
            names = ", ".join(str(s.DisplayName) for s in stores)
            raise OutlookError(f"No mailbox matches {mailboxes}. Available: {names}")
        return picked
    if search_in == "mailbox":
        return [s for s in stores if _safe_get(s, "StoreID") == default_id] or stores[:1]
    if search_in == "archive":
        picked = [s for s in stores if store_kind(s) == "archive"]
        if not picked:
            raise OutlookError(
                "No archive mailbox is open in this Outlook profile. If you have an "
                "online archive it should appear in Outlook's folder pane; see "
                "outlook_list_mailboxes."
            )
        return picked
    if search_in == "all":
        return [s for s in stores if store_kind(s) != "public_folders"]
    raise OutlookError("search_in must be one of: 'folder', 'mailbox', 'archive', 'all'.")


def iter_mail_folders(
    stores: list[Any],
    *,
    include_deleted: bool = False,
    max_depth: int = 8,
    max_folders: int = 1000,
) -> Any:
    """Yield ``(folder, "Mailbox/Path/To/Folder")`` for every mail folder."""
    count = 0
    skip = set(_SKIP_FOLDER_NAMES)
    if not include_deleted:
        skip |= _SKIP_WHEN_EXCLUDING_DELETED

    def walk(folder: Any, path: str, depth: int):
        nonlocal count
        if count >= max_folders:
            return
        if str(_safe_get(folder, "Name", "")).lower() in skip:
            return
        if _safe_get(folder, "DefaultItemType", MAIL_ITEM_TYPE) == MAIL_ITEM_TYPE and depth > 0:
            count += 1
            yield folder, path
        if depth >= max_depth:
            return
        try:
            subs = list(folder.Folders)
        except Exception:
            return
        for sub in subs:
            yield from walk(sub, f"{path}/{sub.Name}", depth + 1)

    for store in stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        yield from walk(root, str(_safe_get(store, "DisplayName", "")), 0)
