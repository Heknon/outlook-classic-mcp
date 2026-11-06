"""Contact COM operations.

Corporate accounts commonly have an *empty* default Contacts folder —
real entries live in sibling contact folders (Recipient Cache, etc.)
and in the directory (Exchange Global Address List). So list/search
cover every contact folder in every store, and search can also scan
the GAL.
"""

from __future__ import annotations

from typing import Any, Iterator

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import OL_CLASS_CONTACT, OL_CONTACT_ITEM


def _contact_summary(c: Any) -> dict[str, Any]:
    return {
        "entry_id": _safe_get(c, "EntryID"),
        "full_name": _safe_get(c, "FullName", ""),
        "email": _safe_get(c, "Email1Address"),
        "company": _safe_get(c, "CompanyName"),
        "job_title": _safe_get(c, "JobTitle"),
        "mobile": _safe_get(c, "MobileTelephoneNumber"),
    }


def _contact_full(c: Any) -> dict[str, Any]:
    return {
        **_contact_summary(c),
        "business_phone": _safe_get(c, "BusinessTelephoneNumber"),
        "home_phone": _safe_get(c, "HomeTelephoneNumber"),
        "address": _safe_get(c, "BusinessAddress"),
        "notes": _safe_get(c, "Body", ""),
    }


def _contact_folders(namespace: Any) -> Iterator[tuple[Any, str]]:
    """Yield ``(folder, path)`` for every contact folder in every store.

    ``DefaultItemType`` uses the same OlItemType enum as CreateItem, so
    contact folders are the ones reporting ``OL_CONTACT_ITEM`` (2).
    """

    def walk(folder: Any, path: str, depth: int) -> Iterator[tuple[Any, str]]:
        if _safe_get(folder, "DefaultItemType") == OL_CONTACT_ITEM:
            yield folder, path
        if depth >= 3:
            return
        try:
            subs = list(folder.Folders)
        except Exception:
            return
        for sub in subs:
            yield from walk(sub, f"{path}/{sub.Name}", depth + 1)

    for store in namespace.Stores:
        try:
            root = store.GetRootFolder()
        except Exception:
            continue
        yield from walk(root, str(store.DisplayName), 0)


def _matches_all(haystack: str, words: list[str]) -> bool:
    return all(w in haystack for w in words)


def list_contacts(outlook: Any, namespace: Any, *, limit: int = 50, offset: int = 0) -> dict[str, Any]:
    results: list[dict[str, Any]] = []
    skipped = 0
    for folder, path in _contact_folders(namespace):
        items = folder.Items
        try:
            items.Sort("[FullName]")
        except Exception:
            pass
        for c in items:
            if _safe_get(c, "Class") != OL_CLASS_CONTACT:
                continue
            if skipped < offset:
                skipped += 1
                continue
            row = _contact_summary(c)
            row["folder"] = path
            results.append(row)
            if len(results) >= limit:
                return {
                    "count": len(results),
                    "offset": offset,
                    "items": results,
                    "has_more": True,
                }
    return {
        "count": len(results),
        "offset": offset,
        "items": results,
        "has_more": False,
    }


def _search_directory(namespace: Any, words: list[str], limit: int) -> list[dict[str, Any]]:
    """Scan the Global Address List for entries matching all words.

    Matching is on the display name; details (SMTP, title, company) are
    fetched only for matches to keep the scan cheap. Large GALs take a
    few seconds — acceptable for an explicit search call.
    """
    out: list[dict[str, Any]] = []
    try:
        entries = namespace.GetGlobalAddressList().AddressEntries
    except Exception:
        return out  # no Exchange directory on this profile
    for entry in entries:
        name = str(_safe_get(entry, "Name", "") or "")
        # Address is cheap to read and, for Exchange entries, is the legacy
        # DN ending in the alias (".../cn=jsmith"), so "jsmith" matches too.
        address = str(_safe_get(entry, "Address", "") or "")
        if not _matches_all(f"{name} {address}".lower(), words):
            continue
        row: dict[str, Any] = {
            "full_name": name,
            "email": None,
            "company": None,
            "job_title": None,
            "mobile": None,
            "source": "directory",
        }
        try:
            exu = entry.GetExchangeUser()
        except Exception:
            exu = None
        if exu is not None:
            row["email"] = _safe_get(exu, "PrimarySmtpAddress")
            row["company"] = _safe_get(exu, "CompanyName")
            row["job_title"] = _safe_get(exu, "JobTitle")
            row["mobile"] = _safe_get(exu, "MobileTelephoneNumber")
            row["department"] = _safe_get(exu, "Department")
            row["office"] = _safe_get(exu, "OfficeLocation")
        out.append(row)
        if len(out) >= limit:
            break
    return out


def search_contacts(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    limit: int = 25,
    include_directory: bool = True,
) -> dict[str, Any]:
    words = [w for w in query.lower().split() if w] or [query.lower()]
    results: list[dict[str, Any]] = []
    seen_emails: set[str] = set()

    for folder, path in _contact_folders(namespace):
        if len(results) >= limit:
            break
        for c in folder.Items:
            if _safe_get(c, "Class") != OL_CLASS_CONTACT:
                continue
            haystack = " ".join(
                str(_safe_get(c, attr, "") or "")
                for attr in ("FullName", "Email1Address", "CompanyName", "JobTitle")
            ).lower()
            if _matches_all(haystack, words):
                row = _contact_summary(c)
                row["folder"] = path
                row["source"] = "contacts"
                results.append(row)
                if row.get("email"):
                    seen_emails.add(str(row["email"]).lower())
                if len(results) >= limit:
                    break

    searched_directory = False
    if include_directory and len(results) < limit:
        searched_directory = True
        for row in _search_directory(namespace, words, limit - len(results)):
            if row.get("email") and str(row["email"]).lower() in seen_emails:
                continue
            results.append(row)

    return {
        "query": query,
        "count": len(results),
        "items": results,
        "searched_directory": searched_directory,
    }


def get_contact(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    return _contact_full(get_item_by_id(namespace, entry_id))


def _entry_details(entry: Any) -> dict[str, Any]:
    smtp = _safe_get(entry, "Address")
    details: dict[str, Any] = {"display_name": _safe_get(entry, "Name")}
    try:
        exu = entry.GetExchangeUser()
    except Exception:
        exu = None
    if exu is not None:
        if _safe_get(exu, "PrimarySmtpAddress"):
            smtp = exu.PrimarySmtpAddress
        details.update(
            {
                "job_title": _safe_get(exu, "JobTitle"),
                "department": _safe_get(exu, "Department"),
                "office": _safe_get(exu, "OfficeLocation"),
                "phone": _safe_get(exu, "BusinessTelephoneNumber"),
                "alias": _safe_get(exu, "Alias"),
            }
        )
    details["smtp_address"] = smtp
    details["legacy_address"] = _safe_get(entry, "Address")
    return {k: v for k, v in details.items() if v not in (None, "")}


def lookup_person(namespace: Any, name: str, max_candidates: int = 10) -> dict[str, Any]:
    """Ctrl+K: resolve a (partial) name, or list the candidates.

    Returns ``{"status": "resolved", "person": {...}}`` for a unique match,
    ``{"status": "ambiguous", "candidates": [...]}`` when several people
    match, or ``{"status": "not_found"}``.
    """
    recipient = namespace.CreateRecipient(name)
    try:
        recipient.Resolve()
    except Exception:
        pass
    if _safe_get(recipient, "Resolved", False):
        return {"status": "resolved", "query": name, "person": _entry_details(recipient.AddressEntry)}

    found = search_contacts(None, namespace, query=name, limit=max_candidates)
    candidates = [
        {
            k: v
            for k, v in {
                "display_name": row.get("full_name"),
                "smtp_address": row.get("email"),
                "job_title": row.get("job_title"),
                "department": row.get("department"),
                "office": row.get("office"),
                "company": row.get("company"),
                "source": row.get("source"),
            }.items()
            if v
        }
        for row in found["items"]
    ]
    candidates = [c for c in candidates if c.get("smtp_address") or c.get("display_name")]
    if len(candidates) == 1 and candidates[0].get("smtp_address"):
        return {"status": "resolved", "query": name, "person": candidates[0]}
    if candidates:
        return {
            "status": "ambiguous",
            "query": name,
            "candidates": candidates,
            "note": "Several people match. Ask the user which one, then use their smtp_address.",
        }
    return {
        "status": "not_found",
        "query": name,
        "note": "Nobody in your contacts or the address book matches. Try another spelling or the email address.",
    }


def resolve_name(outlook: Any, namespace: Any, *, name: str, max_candidates: int = 10) -> dict[str, Any]:
    """Resolve a name like Ctrl+K in Outlook (with candidates when ambiguous)."""
    out = lookup_person(namespace, name, max_candidates)
    # Backwards-compatible fields.
    out["resolved"] = out["status"] == "resolved"
    if out["resolved"]:
        out["display_name"] = out["person"].get("display_name")
        out["smtp_address"] = out["person"].get("smtp_address")
    return out
