"""Mail COM operations."""

from __future__ import annotations

import ntpath
import os
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id, resolve_folder
from outlook_mcp.constants import (
    IMPORTANCE_MAP,
    OL_BCC,
    OL_BY_VALUE,
    OL_CC,
    OL_CLASS_MAIL,
    OL_CLASS_MEETING_REQUEST,
    OL_FOLDER_DRAFTS,
    OL_FORMAT_HTML,
    OL_FORMAT_PLAIN,
    OL_MAIL_ITEM,
    OL_TO,
    PR_ATTACH_CONTENT_ID,
    PR_ATTACHMENT_HIDDEN,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.formatting import from_iso, to_dasl_date, to_iso, truncate
from outlook_mcp.utils.html_compose import (
    EMPTY_PARAGRAPH,
    append_to_body,
    body_insert_position,
    compose_fragment,
    has_auto_signature,
    insert_at_top,
    new_html_document,
    text_to_html,
)
from outlook_mcp.utils.paths import validate_attachment_path, validate_output_dir
from outlook_mcp.utils.policy import DRAFT_ONLY_NOTE, draft_only
from outlook_mcp.utils.safety import safe_dasl
from outlook_mcp.utils.signatures import (
    Signature,
    default_signature_name,
    list_signatures,
    load_signature,
    signatures_dir,
)

WINDOWS_RESERVED_DEVICE_NAMES = {"CON", "PRN", "AUX", "NUL", "CLOCK$"} | {
    f"COM{i}" for i in range(1, 10)
} | {f"LPT{i}" for i in range(1, 10)}

# PR_SENDER_SMTP_ADDRESS — unlike urn:schemas:httpmail:fromemail, this
# holds the real SMTP address even for Exchange senders (whose fromemail
# is an EX:/O=... distinguished name).
SMTP_PROPTAG = "http://schemas.microsoft.com/mapi/proptag/0x5D02001F"

# OlInspectorClose.olSave
OL_INSPECTOR_SAVE = 0


def split_search_words(query: str) -> tuple[str, list[str]]:
    """Split a search query into (anchor, remaining_words), lowercased.

    The anchor is the longest word — the most selective term to push
    down into the DASL Restrict. The remaining words are verified in
    Python per item, because DASL can't reliably AND two LIKEs on the
    same property (verified live: it returns zero rows).
    """
    words = [w.lower() for w in query.split() if w]
    if not words:
        return query.lower(), []
    anchor = max(words, key=len)
    remaining = list(words)
    remaining.remove(anchor)
    return anchor, remaining


def _search_haystack(item: Any, scope: str) -> str:
    if scope == "subject":
        fields = ("Subject",)
    elif scope == "from":
        fields = ("SenderName", "SenderEmailAddress")
    else:  # subject_body
        fields = ("Subject", "Body")
    return " ".join(str(_safe_get(item, f, "") or "") for f in fields).lower()


def _mail_summary(item: Any) -> dict[str, Any]:
    attachments = _safe_get(item, "Attachments")
    return {
        "entry_id": _safe_get(item, "EntryID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "has_attachments": attachments.Count > 0 if attachments else False,
        "importance": _safe_get(item, "Importance"),
        "preview": truncate(_safe_get(item, "Body", ""), 200),
    }


def _attachment_name(att: Any, index: int) -> Any:
    for prop in ("FileName", "DisplayName"):
        value = _safe_get(att, prop)
        if value:
            return value
    return f"(attachment {index}, name unavailable)"


def _mail_full(
    item: Any,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    result = {
        "entry_id": _safe_get(item, "EntryID"),
        "conversation_id": _safe_get(item, "ConversationID"),
        "subject": _safe_get(item, "Subject", ""),
        "from": _safe_get(item, "SenderName"),
        "from_address": _safe_get(item, "SenderEmailAddress"),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
        "received": to_iso(_safe_get(item, "ReceivedTime")),
        "sent": to_iso(_safe_get(item, "SentOn")),
        "unread": bool(_safe_get(item, "UnRead", False)),
        "flagged": _safe_get(item, "FlagStatus") == 2,  # olFlagMarked
        "importance": _safe_get(item, "Importance"),
        "categories": _safe_get(item, "Categories", ""),
    }
    if include_attachments:
        # Enumerating attachments and reading each one's properties is the
        # slowest and most crash-prone part of fetching a mail: some
        # attachment types (embedded .msg items, OLE objects) raise a
        # com_error on FileName/Size (verified live), which used to fail
        # the whole call even though the body was perfectly readable.
        # _attachment_name()/_safe_get() below fall back instead of raising.
        attachments = []
        raw_attachments = _safe_get(item, "Attachments")
        if raw_attachments:
            for i, att in enumerate(raw_attachments, start=1):
                attachments.append(
                    {
                        "index": i,
                        "filename": _attachment_name(att, i),
                        "size_bytes": _safe_get(att, "Size"),
                    }
                )
        result["attachments"] = attachments
        result["attachment_count"] = len(attachments)
    else:
        # Skip Attachments entirely (not even .Count) so callers that
        # don't need attachments (bulk summarizing/searching) never pay
        # for the riskiest COM traffic in this call.
        result["attachments"] = None
    if include_body:
        body = _safe_get(item, "Body", "") or ""
        if max_body_chars and len(body) > max_body_chars:
            result["body"] = body[:max_body_chars].rstrip()
            result["body_truncated"] = True
            result["body_total_chars"] = len(body)
        else:
            result["body"] = body
    if include_html:
        # Full HTML of a styled corporate mail easily runs to tens of
        # kilobytes — only fetch when explicitly asked for.
        result["html_body"] = _safe_get(item, "HTMLBody", "")
    return result


def list_mails(
    outlook: Any,
    namespace: Any,
    *,
    folder: str | None = "inbox",
    limit: int = 25,
    offset: int = 0,
    unread_only: bool = False,
    since: str | None = None,
    until: str | None = None,
    from_address: str | None = None,
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    clauses: list[str] = []
    if unread_only:
        clauses.append("[UnRead] = True")
    since_dt = from_iso(since)
    until_dt = from_iso(until)
    # DASL property query (locale-independent ISO dates) instead of a plain
    # "[ReceivedTime] >= 'mm/dd/yyyy ...'" Jet literal, which Outlook parses
    # using the Windows regional short-date order — on a non-US locale that
    # silently swaps day/month and matches the wrong emails (verified live).
    if since_dt:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" >= '{to_dasl_date(since_dt)}'")
    if until_dt:
        clauses.append(f"\"urn:schemas:httpmail:datereceived\" <= '{to_dasl_date(until_dt)}'")

    if clauses:
        # unread_only's plain "[UnRead] = True" clause can't share an @SQL=
        # DASL restrict with the datereceived clauses above, so split them
        # into two Restrict() calls instead of one mixed-syntax string.
        plain = [c for c in clauses if not c.startswith('"urn:')]
        dasl = [c for c in clauses if c.startswith('"urn:')]
        if plain:
            items = items.Restrict(" AND ".join(plain))
        if dasl:
            items = items.Restrict("@SQL=" + " AND ".join(dasl))

    from_lower = from_address.lower() if from_address else None
    results: list[dict[str, Any]] = []
    skipped = 0
    for item in items:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        if from_lower:
            sender = (_safe_get(item, "SenderEmailAddress") or "").lower()
            if from_lower not in sender:
                continue
        if skipped < offset:
            skipped += 1
            continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "folder": f.Name,
        "count": len(results),
        "offset": offset,
        "limit": limit,
        "items": results,
        "has_more": len(results) == limit,
        "next_offset": offset + len(results) if len(results) == limit else None,
    }


def search_mails(
    outlook: Any,
    namespace: Any,
    *,
    query: str,
    folder: str | None = "inbox",
    limit: int = 25,
    scope: str = "subject_body",
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)

    # Multi-word queries: Restrict on the most selective word, then
    # require the remaining words per item in Python (see
    # split_search_words for why DASL can't do the AND itself).
    remaining: list[str] = []
    if scope == "dasl":
        # Caller is explicitly passing a raw DASL filter; don't mangle it.
        filtered = items.Restrict(query)
    else:
        anchor, remaining = split_search_words(query)
        esc = safe_dasl(anchor)
        if scope == "subject":
            filtered = items.Restrict(
                f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{esc}%'"
            )
        elif scope == "from":
            filtered = items.Restrict(
                f"@SQL=(\"urn:schemas:httpmail:fromemail\" LIKE '%{esc}%' OR "
                f"\"urn:schemas:httpmail:fromname\" LIKE '%{esc}%' OR "
                f"\"{SMTP_PROPTAG}\" LIKE '%{esc}%')"
            )
        else:  # subject_body
            filtered = items.Restrict(
                f"@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%{esc}%' OR "
                f"\"urn:schemas:httpmail:textdescription\" LIKE '%{esc}%')"
            )

    results: list[dict[str, Any]] = []
    for item in filtered:
        cls = _safe_get(item, "Class")
        if cls not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        if remaining:
            haystack = _search_haystack(item, scope)
            if not all(word in haystack for word in remaining):
                continue
        results.append(_mail_summary(item))
        if len(results) >= limit:
            break

    return {
        "query": query,
        "scope": scope,
        "folder": f.Name,
        "count": len(results),
        "items": results,
    }


def get_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    include_body: bool = True,
    include_html: bool = False,
    max_body_chars: int = 10000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    return _mail_full(
        get_item_by_id(namespace, entry_id),
        include_body=include_body,
        include_html=include_html,
        max_body_chars=max_body_chars,
        include_attachments=include_attachments,
    )


def find_account(outlook: Any, account: str) -> Any:
    """Return the Outlook Account whose SMTP address or display name matches.

    Matching is case-insensitive. Raises OutlookError listing the available
    accounts if nothing matches.
    """
    wanted = account.strip().lower()
    accounts = outlook.Session.Accounts
    available: list[str] = []
    for i in range(1, accounts.Count + 1):
        acc = accounts.Item(i)
        smtp = str(_safe_get(acc, "SmtpAddress", "") or "")
        display = str(_safe_get(acc, "DisplayName", "") or "")
        available.append(smtp or display)
        if wanted in (smtp.lower(), display.lower()):
            return acc
    raise OutlookError(
        f"Account '{account}' not found. Available accounts: {', '.join(available) or 'none'}"
    )


# --------------------------------------------------------------------------
# Composition helpers shared by send / reply / forward
# --------------------------------------------------------------------------


def _default_account(namespace: Any) -> Any:
    """The account that owns the default store, else the first account."""
    try:
        accounts = namespace.Accounts
        default_store_id = _safe_get(_safe_get(namespace, "DefaultStore"), "StoreID")
        first = None
        for i in range(1, accounts.Count + 1):
            acc = accounts.Item(i)
            if first is None:
                first = acc
            store = _safe_get(acc, "DeliveryStore")
            if default_store_id and _safe_get(store, "StoreID") == default_store_id:
                return acc
        return first
    except Exception:
        return None


def _account_identifiers(acc: Any) -> list[str]:
    if acc is None:
        return []
    return [
        str(v)
        for v in (_safe_get(acc, "DisplayName"), _safe_get(acc, "SmtpAddress"))
        if v
    ]


def _wants_default_signature(signature: Any) -> bool:
    return signature is True or (
        isinstance(signature, str) and signature.strip().lower() in ("default", "true")
    )


def _resolve_signature(
    namespace: Any, signature: Any, kind: str, account: Any
) -> Signature | None:
    """Load the requested signature, or None if none applies.

    ``signature`` is falsy (no signature), True / "default" (the account's
    default for ``kind`` — 'new' or 'reply'), or a signature name.
    """
    if not signature:
        return None
    if _wants_default_signature(signature):
        name = default_signature_name(
            kind,
            _safe_get(namespace, "CurrentProfileName"),
            _account_identifiers(account),
        )
        return load_signature(name) if name else None
    return load_signature(str(signature))


def _attach_inline_images(item: Any, sig: Signature | None) -> None:
    """Attach signature images as hidden inline attachments (cid:)."""
    if sig is None:
        return
    for path, cid in sig.images:
        att = item.Attachments.Add(path, OL_BY_VALUE)
        pa = att.PropertyAccessor
        pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
        try:
            pa.SetProperty(PR_ATTACHMENT_HIDDEN, True)
        except Exception:  # not fatal: the image just also shows as a file
            pass


def _add_recipients(item: Any, addresses: list[str] | None, kind: int) -> None:
    for addr in addresses or []:
        if addr and addr.strip():
            rec = item.Recipients.Add(addr.strip())
            rec.Type = kind


def _apply_importance(item: Any, importance: str | None) -> None:
    if importance:
        code = IMPORTANCE_MAP.get(importance.lower())
        if code is None:
            raise OutlookError("importance must be one of: 'low', 'normal', 'high'.")
        item.Importance = code


def _write_reply_body(
    item: Any, body: str, html: bool, sig: Signature | None
) -> None:
    """Put new text (and signature) above the quoted history of a reply/forward.

    For HTML items the text goes inside the existing ``<body>`` /
    ``WordSection1`` so it picks up the message's own styles; the quoted
    history and its formatting are left untouched. Rich Text items are
    converted to HTML first. Plain-text items get plain-text treatment
    unless the caller asked for HTML or a signature.
    """
    if not body and sig is None:
        return
    body_format = _safe_get(item, "BodyFormat")
    if body_format == OL_FORMAT_PLAIN and not html and sig is None:
        item.Body = body + "\n\n" + (_safe_get(item, "Body", "") or "")
        return
    # HTML, or Rich Text (common for internal Exchange mail): assigning .Body
    # would throw away the formatting of the quoted thread, so work in HTML.
    if body_format != OL_FORMAT_HTML:
        item.BodyFormat = OL_FORMAT_HTML  # Outlook converts the existing body
    fragment = compose_fragment(
        body,
        is_html=html,
        signature_html=sig.html if sig else None,
        trailing_gap=True,
    )
    item.HTMLBody = insert_at_top(_safe_get(item, "HTMLBody", "") or "", fragment)


def _write_new_body(item: Any, body: str, html: bool, sig: Signature | None) -> None:
    if sig is None:
        if html:
            item.BodyFormat = OL_FORMAT_HTML
            item.HTMLBody = body
        else:
            item.BodyFormat = OL_FORMAT_PLAIN
            item.Body = body
        return
    item.BodyFormat = OL_FORMAT_HTML
    if html and body_insert_position(body) is not None:
        # Caller supplied a full HTML document: keep its <head>/styles and
        # put the signature at the end of its body.
        item.HTMLBody = append_to_body(body, EMPTY_PARAGRAPH + sig.html)
        return
    fragment = compose_fragment(body, is_html=html, signature_html=sig.html)
    item.HTMLBody = new_html_document(fragment)


def _insert_default_signature_via_inspector(
    namespace: Any, mail: Any, body: str, html: bool
) -> Any:
    """Fallback when no default signature is found in the registry.

    Outlook inserts the account's default signature when an inspector is
    created for a new item. The hidden inspector must be closed and the
    item reloaded before it can be sent (Send() fails with E_INVALIDARG
    otherwise). Returns the (possibly reloaded) item.
    """
    mail.BodyFormat = OL_FORMAT_HTML
    inspector = mail.GetInspector
    existing = mail.HTMLBody or ""
    new_html = body if html else text_to_html(body)
    pos = body_insert_position(existing)
    if pos is not None:
        mail.HTMLBody = existing[:pos] + new_html + EMPTY_PARAGRAPH + existing[pos:]
    else:
        mail.HTMLBody = new_html + EMPTY_PARAGRAPH + existing
    mail.Save()
    entry_id, store_id = mail.EntryID, mail.Parent.StoreID
    inspector.Close(OL_INSPECTOR_SAVE)
    return get_item_by_id(namespace, entry_id, store_id)


def _unresolved(item: Any) -> list[str]:
    """Resolve all recipients; return the names that could not be resolved."""
    try:
        if item.Recipients.ResolveAll():
            return []
    except Exception:
        pass
    names = []
    for r in item.Recipients:
        if not _safe_get(r, "Resolved", True):
            names.append(str(_safe_get(r, "Name", "") or _safe_get(r, "Address", "")))
    return names


def _save_or_send(
    item: Any,
    *,
    save_only: bool,
    result: dict[str, Any],
    notes: list[str],
) -> dict[str, Any]:
    """Shared tail of every compose tool: resolve, then save or send.

    Recipients are resolved before saving too: a draft saved with
    unresolved recipients has no address on them, and new Outlook / OWA
    then fail to send it after an edit.
    """
    unresolved = _unresolved(item)
    forced = draft_only() and not save_only
    if save_only or forced:
        item.Save()
        if forced:
            notes.append(DRAFT_ONLY_NOTE)
        if unresolved:
            notes.append(
                "Some recipients could not be resolved and need fixing before "
                f"sending: {', '.join(unresolved)}"
            )
        out = {
            "status": "saved_to_drafts",
            "entry_id": _safe_get(item, "EntryID"),
            "subject": _safe_get(item, "Subject", ""),
            **result,
        }
    else:
        if unresolved:
            raise OutlookError(
                "Not sent: these recipients could not be resolved: "
                f"{', '.join(unresolved)}. Fix the addresses (outlook_resolve_name "
                "can help) or save as a draft with save_only=true."
            )
        # Cache properties BEFORE Send(): once sent, the COM object has moved
        # to Sent Items and reading it raises "item has been moved or
        # deleted" — surfacing that as a failure makes agents retry and send
        # duplicates even though Send() succeeded.
        subject = _safe_get(item, "Subject", "")
        item.Send()
        out = {"status": "sent", "subject": subject, **result}
    if notes:
        out["notes"] = notes
    return out


# --------------------------------------------------------------------------
# Compose tools
# --------------------------------------------------------------------------


def send_mail(
    outlook: Any,
    namespace: Any,
    *,
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    importance: str = "normal",
    save_only: bool = False,
    account: str | None = None,
    signature: Any = None,
) -> dict[str, Any]:
    # Validate paths before creating anything, so a bad path can't leave a
    # half-built item behind.
    attachment_paths = [validate_attachment_path(p) for p in attachments or []]

    acc = None
    if account:
        acc = find_account(outlook, account)
        # Create the item inside the account's own Drafts folder so that a
        # saved draft lands in that mailbox, not in the default profile's.
        drafts = acc.DeliveryStore.GetDefaultFolder(OL_FOLDER_DRAFTS)
        mail = drafts.Items.Add(OL_MAIL_ITEM)
        mail.SendUsingAccount = acc
    else:
        mail = outlook.CreateItem(OL_MAIL_ITEM)

    mail.To = "; ".join(to)
    if cc:
        mail.CC = "; ".join(cc)
    if bcc:
        mail.BCC = "; ".join(bcc)
    mail.Subject = subject
    _apply_importance(mail, importance)

    notes: list[str] = []
    sig = _resolve_signature(namespace, signature, "new", acc or _default_account(namespace))
    if signature and sig is None:
        # "default" requested but no default is configured in the registry:
        # let Outlook insert it the way it does for a new message window.
        mail = _insert_default_signature_via_inspector(namespace, mail, body, html)
    else:
        _attach_inline_images(mail, sig)
        _write_new_body(mail, body, html, sig)

    for path in attachment_paths:
        mail.Attachments.Add(path)

    result: dict[str, Any] = {"to": to, "cc": cc or [], "bcc": bcc or []}
    if sig is not None:
        result["signature"] = sig.name
    return _save_or_send(mail, save_only=save_only, result=result, notes=notes)


def _reply_like(
    namespace: Any,
    item: Any,
    *,
    kind_label: str,
    source_entry_id: str,
    body: str,
    html: bool,
    to: list[str] | None,
    cc: list[str] | None,
    bcc: list[str] | None,
    subject: str | None,
    importance: str | None,
    attachments: list[str] | None,
    signature: Any,
    save_only: bool,
    extra: dict[str, Any],
) -> dict[str, Any]:
    attachment_paths = [validate_attachment_path(p) for p in attachments or []]
    notes: list[str] = []

    _add_recipients(item, to, OL_TO)
    _add_recipients(item, cc, OL_CC)
    _add_recipients(item, bcc, OL_BCC)
    if subject:
        item.Subject = subject
    _apply_importance(item, importance)

    sig = None
    if signature:
        if _wants_default_signature(signature) and has_auto_signature(
            _safe_get(item, "HTMLBody", "")
        ):
            notes.append("Outlook already added your signature; not added twice.")
        else:
            account = _safe_get(item, "SendUsingAccount") or _default_account(namespace)
            sig = _resolve_signature(namespace, signature, "reply", account)
            if sig is None:
                notes.append(
                    "No default reply/forward signature is configured in Outlook; "
                    "pass a signature name (see outlook_list_signatures)."
                )
    _attach_inline_images(item, sig)
    _write_reply_body(item, body, html, sig)

    for path in attachment_paths:
        item.Attachments.Add(path)

    result = {kind_label: source_entry_id, **extra}
    if sig is not None:
        result["signature"] = sig.name
    if not _safe_get(_safe_get(item, "Recipients"), "Count", 0):
        if not (save_only or draft_only()):
            raise OutlookError(
                "This message has no recipients. Pass `to`, or save it as a draft "
                "with save_only=true and add recipients in Outlook."
            )
    return _save_or_send(item, save_only=save_only, result=result, notes=notes)


def reply_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    body: str,
    reply_all: bool = False,
    html: bool = False,
    attachments: list[str] | None = None,
    to: list[str] | None = None,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    subject: str | None = None,
    importance: str | None = None,
    signature: Any = None,
    save_only: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    reply = original.ReplyAll() if reply_all else original.Reply()
    return _reply_like(
        namespace,
        reply,
        kind_label="in_reply_to",
        source_entry_id=entry_id,
        body=body,
        html=html,
        to=to,
        cc=cc,
        bcc=bcc,
        subject=subject,
        importance=importance,
        attachments=attachments,
        signature=signature,
        save_only=save_only,
        extra={"reply_all": reply_all},
    )


def forward_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    to: list[str] | None = None,
    body: str = "",
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    html: bool = False,
    attachments: list[str] | None = None,
    subject: str | None = None,
    importance: str | None = None,
    signature: Any = None,
    save_only: bool = False,
) -> dict[str, Any]:
    original = get_item_by_id(namespace, entry_id)
    fwd = original.Forward()
    return _reply_like(
        namespace,
        fwd,
        kind_label="forwarded",
        source_entry_id=entry_id,
        body=body,
        html=html,
        to=to,
        cc=cc,
        bcc=bcc,
        subject=subject,
        importance=importance,
        attachments=attachments,
        signature=signature,
        save_only=save_only,
        extra={"to": to or []},
    )


def send_draft(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    """Send a mail that is already saved as a draft (e.g. after review)."""
    if draft_only():
        raise OutlookError(
            "Sending is disabled: draft-only mode is on (OUTLOOK_MCP_DRAFT_ONLY). "
            "Send the draft from Outlook."
        )
    item = get_item_by_id(namespace, entry_id)
    if _safe_get(item, "Class") != OL_CLASS_MAIL:
        raise OutlookError("That item is not a mail message.")
    if _safe_get(item, "Sent", False):
        raise OutlookError("That message has already been sent; it is not a draft.")
    if not _safe_get(_safe_get(item, "Recipients"), "Count", 0):
        raise OutlookError("The draft has no recipients. Add them in Outlook first.")
    unresolved = _unresolved(item)
    if unresolved:
        raise OutlookError(
            "Not sent: these recipients could not be resolved: "
            f"{', '.join(unresolved)}. Fix them in Outlook first."
        )
    result = {
        "status": "sent",
        "subject": _safe_get(item, "Subject", ""),
        "to": _safe_get(item, "To", ""),
        "cc": _safe_get(item, "CC", ""),
        "bcc": _safe_get(item, "BCC", ""),
    }
    item.Send()
    return result


def list_signatures_info(outlook: Any, namespace: Any) -> dict[str, Any]:
    acc = _default_account(namespace)
    profile = _safe_get(namespace, "CurrentProfileName")
    ids = _account_identifiers(acc)
    return {
        "folder": str(signatures_dir()),
        "signatures": list_signatures(),
        "default_new": default_signature_name("new", profile, ids),
        "default_reply_forward": default_signature_name("reply", profile, ids),
    }


def move_mail(outlook: Any, namespace: Any, *, entry_id: str, target_folder: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    target = resolve_folder(namespace, target_folder)
    moved = item.Move(target)
    return {"status": "moved", "new_entry_id": moved.EntryID, "folder": target.Name}


def delete_mail(outlook: Any, namespace: Any, *, entry_id: str) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    subject = _safe_get(item, "Subject", "")
    item.Delete()
    return {"status": "deleted", "subject": subject, "entry_id": entry_id}


def mark_mail(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    read: bool | None = None,
    flagged: bool | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    if read is not None:
        item.UnRead = not read
    if flagged is not None:
        item.FlagStatus = 2 if flagged else 0
    item.Save()
    return {
        "status": "updated",
        "entry_id": entry_id,
        "unread": bool(item.UnRead),
        "flagged": item.FlagStatus == 2,
    }


def save_attachments(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    output_dir: str,
    attachment_index: int | None = None,
) -> dict[str, Any]:
    item = get_item_by_id(namespace, entry_id)
    out_dir = validate_output_dir(output_dir)
    saved: list[str] = []
    attachments = list(item.Attachments)
    if attachment_index is not None:
        if attachment_index < 1 or attachment_index > len(attachments):
            raise OutlookError(
                f"attachment_index {attachment_index} out of range "
                f"(message has {len(attachments)} attachments, 1-indexed)."
            )
        attachments = [attachments[attachment_index - 1]]

    for att in attachments:
        # Sender-controlled filename. Reject anything containing path
        # separators, drive-letter prefixes, dot-only names, or reserved
        # Windows device names. These are signals of a path-traversal
        # attempt by the sender, not legitimate attachments.
        raw = att.FileName or ""
        if not raw or raw in (".", ".."):
            raise OutlookError(f"Attachment has invalid filename: {raw!r}")
        if "\\" in raw or "/" in raw:
            raise OutlookError(
                f"Attachment filename contains path separators "
                f"(rejected for safety): {raw!r}"
            )
        if ":" in raw:
            raise OutlookError(
                f"Attachment filename contains colon "
                f"(rejected for safety): {raw!r}"
            )
        # Defense in depth: basename should be a no-op after the checks
        # above, but use it anyway in case ntpath sees something we missed.
        safe_name = ntpath.basename(raw)
        if safe_name != raw:
            raise OutlookError(
                f"Attachment filename did not normalize cleanly: {raw!r}"
            )
        stem = safe_name.lstrip(".").split(".", 1)[0].upper()
        if stem in WINDOWS_RESERVED_DEVICE_NAMES:
            raise OutlookError(
                f"Attachment has reserved Windows device name: {safe_name!r}"
            )

        # Mails often carry several attachments with the same name (e.g.
        # multiple inline "image.png") — uniquify instead of overwriting.
        target = os.path.join(out_dir, safe_name)
        base, ext = os.path.splitext(safe_name)
        counter = 1
        while os.path.exists(target):
            target = os.path.join(out_dir, f"{base} ({counter}){ext}")
            counter += 1
        att.SaveAsFile(target)
        saved.append(target)
    return {
        "status": "saved",
        "count": len(saved),
        "files": saved,
        "output_dir": out_dir,
    }
