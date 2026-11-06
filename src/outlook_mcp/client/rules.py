"""Mail-rule COM operations: list, explain, create, delete, run, toggle.

Rules live on the default store: ``Store.GetRules()`` returns the whole
collection and ``Rules.Save()`` writes the whole collection back to
Exchange (rules are stored server-side; a few combinations run only in
Outlook — ``Rule.IsLocalRule``). Because Save rewrites everything, each
operation fetches the collection fresh right before changing it.

Creating and running rules is two-step: without ``confirm=True`` the
tools return a preview (the rule in plain terms plus the recent mails it
would match) and change nothing.

Safety: forward / redirect actions send mail automatically, so they are
refused in draft-only mode (the default) and, even when sending is
allowed, limited to the user's own email domain unless
``OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD=1``. Extra internal domains can be
listed in ``OUTLOOK_MCP_INTERNAL_DOMAINS`` (comma-separated).
"""

from __future__ import annotations

import datetime as dt
import os
from typing import Any

from outlook_mcp.client.folders import _safe_get, resolve_folder
from outlook_mcp.constants import OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import RuleActions, RuleConditions
from outlook_mcp.utils.formatting import to_iso
from outlook_mcp.utils.policy import draft_only

OL_RULE_RECEIVE = 0
_EXECUTE_OPTION = {"all": 0, "read": 1, "unread": 2}  # OlRuleExecuteOption
_MARK_INTERVAL = {"today": 0, "tomorrow": 1, "this_week": 2, "next_week": 3, "no_date": 4}
_IMPORTANCE = {"low": 0, "normal": 1, "high": 2}
_IMPORTANCE_NAMES = {v: k for k, v in _IMPORTANCE.items()}

# Condition field -> (RuleConditions COM property, kind, value property)
_CONDITIONS = {
    "from_people": ("From", "recipients", None),
    "sent_to": ("SentTo", "recipients", None),
    "subject_words": ("Subject", "words", "Text"),
    "body_words": ("Body", "words", "Text"),
    "subject_or_body_words": ("BodyOrSubject", "words", "Text"),
    "sender_address_words": ("SenderAddress", "words", "Address"),
    "recipient_address_words": ("RecipientAddress", "words", "Address"),
    "header_words": ("MessageHeader", "words", "Text"),
    "categories": ("Category", "words", "Categories"),
    "importance": ("Importance", "importance", "Importance"),
    "to_me": ("ToMe", "flag", None),
    "only_to_me": ("OnlyToMe", "flag", None),
    "cc_me": ("CC", "flag", None),
    "to_or_cc_me": ("ToOrCc", "flag", None),
    "has_attachment": ("HasAttachment", "flag", None),
}

# Action field -> (RuleActions COM property, kind)
_ACTIONS = {
    "move_to_folder": ("MoveToFolder", "folder"),
    "copy_to_folder": ("CopyToFolder", "folder"),
    "categories": ("AssignToCategory", "categories"),
    "clear_categories": ("ClearCategories", "flag"),
    "mark_as_task": ("MarkAsTask", "mark"),
    "delete": ("Delete", "flag"),
    "delete_permanently": ("DeletePermanently", "flag"),
    "stop_processing": ("Stop", "flag"),
    "desktop_alert": ("DesktopAlert", "flag"),
    "forward_to": ("Forward", "recipients"),
    "forward_as_attachment_to": ("ForwardAsAttachment", "recipients"),
    "redirect_to": ("Redirect", "recipients"),
}
_SENDING_ACTIONS = ("forward_to", "forward_as_attachment_to", "redirect_to")


# ------------------------------------------------------------------ reading


def _recipient_names(obj: Any) -> list[str]:
    out = []
    recips = _safe_get(obj, "Recipients")
    if recips is None:
        return out
    for r in recips:
        smtp = None
        try:
            smtp = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except Exception:
            smtp = None
        out.append(str(smtp or _safe_get(r, "Address") or _safe_get(r, "Name") or ""))
    return [x for x in out if x]


def _read_side(container: Any, mapping: dict, kinds: str) -> dict[str, Any]:
    out: dict[str, Any] = {}
    if container is None:
        return out
    for field, spec in mapping.items():
        prop, kind = spec[0], spec[1]
        obj = _safe_get(container, prop)
        if obj is None or not _safe_get(obj, "Enabled", False):
            continue
        if kind == "flag":
            out[field] = True
        elif kind == "recipients":
            out[field] = _recipient_names(obj)
        elif kind == "words":
            out[field] = list(_safe_get(obj, spec[2], ()) or ())
        elif kind == "importance":
            out[field] = _IMPORTANCE_NAMES.get(_safe_get(obj, "Importance"), "normal")
        elif kind == "folder":
            folder = _safe_get(obj, "Folder")
            out[field] = str(_safe_get(folder, "FolderPath") or _safe_get(folder, "Name") or "")
        elif kind == "categories":
            out[field] = list(_safe_get(obj, "Categories", ()) or ())
        elif kind == "mark":
            code = _safe_get(obj, "MarkInterval", 4)
            out[field] = next((k for k, v in _MARK_INTERVAL.items() if v == code), "no_date")
    return out


def describe_rule(rule: Any, index: int) -> dict[str, Any]:
    return {
        "index": index,
        "name": _safe_get(rule, "Name"),
        "enabled": bool(_safe_get(rule, "Enabled", False)),
        "runs_on": "outgoing" if _safe_get(rule, "RuleType", 0) == 1 else "incoming",
        "client_only": bool(_safe_get(rule, "IsLocalRule", False)),
        "conditions": _read_side(_safe_get(rule, "Conditions"), _CONDITIONS, "conditions"),
        "exceptions": _read_side(_safe_get(rule, "Exceptions"), _CONDITIONS, "conditions"),
        "actions": _read_side(_safe_get(rule, "Actions"), _ACTIONS, "actions"),
    }


def _rules(namespace: Any) -> Any:
    return namespace.DefaultStore.GetRules()


def _find_rule(rules: Any, name: str) -> tuple[int, Any] | None:
    wanted = name.strip().lower()
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        if str(_safe_get(rule, "Name", "")).strip().lower() == wanted:
            return i + 1, rule
    return None


def list_rules(outlook: Any, namespace: Any, *, details: bool = False) -> dict[str, Any]:
    rules = _rules(namespace)
    items = []
    for i in range(rules.Count):
        rule = rules.Item(i + 1)
        if details:
            items.append(describe_rule(rule, i + 1))
        else:
            items.append({"index": i + 1, "name": rule.Name, "enabled": bool(rule.Enabled)})
    return {"count": len(items), "items": items}


def toggle_rule(outlook: Any, namespace: Any, *, rule_name: str, enabled: bool) -> dict[str, Any]:
    rules = _rules(namespace)
    found = _find_rule(rules, rule_name)
    if found is None:
        raise OutlookError(f"Rule '{rule_name}' not found. Use outlook_list_rules to see available rules.")
    found[1].Enabled = enabled
    rules.Save()
    return {"status": "updated", "rule": found[1].Name, "enabled": bool(enabled)}


# ------------------------------------------------------------------ safety


def _my_smtp(namespace: Any) -> str:
    me = _safe_get(namespace, "CurrentUser")
    try:
        return str(me.AddressEntry.GetExchangeUser().PrimarySmtpAddress or "")
    except Exception:
        return str(_safe_get(me, "Address", "") or "")


def _internal_domains(namespace: Any) -> set[str]:
    domains = set()
    mine = _my_smtp(namespace)
    if "@" in mine:
        domains.add(mine.rsplit("@", 1)[1].lower())
    for d in os.environ.get("OUTLOOK_MCP_INTERNAL_DOMAINS", "").split(","):
        if d.strip():
            domains.add(d.strip().lstrip("@").lower())
    return domains


def _resolve_smtp(namespace: Any, who: str) -> str | None:
    from outlook_mcp.client.contacts import lookup_person

    if "@" in who and " " not in who.strip():
        return who.strip()
    found = lookup_person(namespace, who)
    if found["status"] == "resolved":
        return found["person"].get("smtp_address")
    return None


def _check_sending_actions(namespace: Any, actions: RuleActions) -> list[str]:
    targets = [a for f in _SENDING_ACTIONS for a in (getattr(actions, f) or [])]
    if not targets:
        return []
    if draft_only():
        raise OutlookError(
            "Refused: forward/redirect rules send mail automatically, and the server is in "
            "draft-only mode (the default). The user can create this rule in Outlook "
            "(File > Manage Rules & Alerts), or enable sending with OUTLOOK_MCP_ALLOW_SEND=1."
        )
    if os.environ.get("OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD", "").strip().lower() in ("1", "true", "yes"):
        return []
    internal = _internal_domains(namespace)
    bad = []
    for who in targets:
        smtp = _resolve_smtp(namespace, who)
        domain = smtp.rsplit("@", 1)[1].lower() if smtp and "@" in smtp else None
        if domain is None or domain not in internal:
            bad.append(f"{who} ({smtp or 'unresolved'})")
    if bad:
        raise OutlookError(
            "Refused: forwarding rules may only send to your own organization "
            f"({', '.join(sorted(internal)) or 'unknown domain'}). Not allowed: {', '.join(bad)}. "
            "Set OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD=1 (or OUTLOOK_MCP_INTERNAL_DOMAINS) if this is intended."
        )
    return []


# ----------------------------------------------------------------- preview


def _me(namespace: Any) -> tuple[str, str]:
    me = _safe_get(namespace, "CurrentUser")
    return str(_safe_get(me, "Name", "") or "").lower(), _my_smtp(namespace).lower()


def _sender(item: Any) -> str:
    return " ".join(
        str(_safe_get(item, p, "") or "")
        for p in ("SenderName", "SenderEmailAddress", "_sender_smtp")
    ).lower()


def _has_any(haystack: str, words: list[str]) -> bool:
    h = haystack.lower()
    return any(w.lower() in h for w in words if w)


def _matches(item: Any, c: RuleConditions, me: tuple[str, str], people: dict[str, list[str]]) -> bool | None:
    """Evaluate conditions like Outlook does: all set conditions must hold;
    word lists match if any word appears. Returns None if a condition
    can't be evaluated from the item (header words)."""
    name, smtp = me
    to = str(_safe_get(item, "To", "") or "").lower()
    cc = str(_safe_get(item, "CC", "") or "").lower()
    checks = []
    if c.from_people:
        s = _sender(item)
        checks.append(any(any(k in s for k in keys) for keys in people["from_people"]))
    if c.sent_to:
        checks.append(any(any(k in to + " " + cc for k in keys) for keys in people["sent_to"]))
    if c.subject_words:
        checks.append(_has_any(str(_safe_get(item, "Subject", "") or ""), c.subject_words))
    if c.body_words:
        checks.append(_has_any(str(_safe_get(item, "Body", "") or ""), c.body_words))
    if c.subject_or_body_words:
        checks.append(_has_any(f"{_safe_get(item, 'Subject', '')} {_safe_get(item, 'Body', '')}", c.subject_or_body_words))
    if c.sender_address_words:
        checks.append(_has_any(_sender(item), c.sender_address_words))
    if c.recipient_address_words:
        checks.append(_has_any(to + " " + cc, c.recipient_address_words))
    if c.categories:
        checks.append(_has_any(str(_safe_get(item, "Categories", "") or ""), c.categories))
    if c.importance:
        checks.append(_safe_get(item, "Importance") == _IMPORTANCE[c.importance])
    if c.has_attachment:
        checks.append((_safe_get(_safe_get(item, "Attachments"), "Count", 0) or 0) > 0)
    mine = [k for k in (name, smtp) if k]
    if c.to_me:
        checks.append(any(k in to for k in mine))
    if c.only_to_me:
        checks.append(any(k in to for k in mine) and ";" not in to and not cc)
    if c.cc_me:
        checks.append(any(k in cc for k in mine))
    if c.to_or_cc_me:
        checks.append(any(k in to + " " + cc for k in mine))
    if c.header_words:
        return None
    return all(checks)


def _people_keys(namespace: Any, people: list[str]) -> list[list[str]]:
    """Match keys per person: name and SMTP address (resolved if possible)."""
    out = []
    from outlook_mcp.client.contacts import lookup_person

    for p in people or []:
        keys = {p.strip().lower()}
        found = lookup_person(namespace, p)
        if found["status"] == "resolved":
            for k in ("display_name", "smtp_address", "legacy_address"):
                v = found["person"].get(k)
                if v:
                    keys.add(str(v).lower())
        out.append(sorted(keys))
    return out


def preview_matches(
    namespace: Any,
    conditions: RuleConditions,
    exceptions: RuleConditions | None,
    *,
    folder: str = "inbox",
    days: int = 30,
    max_scan: int = 500,
    sample: int = 10,
) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    items = f.Items
    items.Sort("[ReceivedTime]", True)
    me = _me(namespace)
    people = {
        "from_people": _people_keys(namespace, conditions.from_people),
        "sent_to": _people_keys(namespace, conditions.sent_to),
    }
    ex_people = {
        "from_people": _people_keys(namespace, exceptions.from_people) if exceptions else [],
        "sent_to": _people_keys(namespace, exceptions.sent_to) if exceptions else [],
    }
    cutoff = dt.datetime.now() - dt.timedelta(days=days)
    scanned = matched = 0
    unknown = False
    samples: list[dict[str, Any]] = []
    for item in items:
        if _safe_get(item, "Class") not in (OL_CLASS_MAIL, OL_CLASS_MEETING_REQUEST):
            continue
        rt = _safe_get(item, "ReceivedTime")
        if rt is not None and hasattr(rt, "year"):
            local = dt.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute)
            if local < cutoff:
                break
        scanned += 1
        if scanned > max_scan:
            break
        hit = _matches(item, conditions, me, people)
        if hit is None:
            unknown = True
            continue
        if hit and exceptions is not None and exceptions.model_dump(exclude_defaults=True):
            ex = _matches(item, exceptions, me, ex_people)
            if ex:
                hit = False
        if hit:
            matched += 1
            if len(samples) < sample:
                samples.append({
                    "subject": _safe_get(item, "Subject", ""),
                    "from": _safe_get(item, "SenderName"),
                    "received": to_iso(rt),
                })
    out: dict[str, Any] = {
        "folder": _safe_get(f, "Name"),
        "scanned": min(scanned, max_scan),
        "days": days,
        "would_match": matched,
        "examples": samples,
    }
    if unknown:
        out["note"] = "Header-word conditions can't be previewed; the match count may be incomplete."
    return out


# ---------------------------------------------------------------- building


def _set_words(obj: Any, prop: str, words: list[str]) -> None:
    setattr(obj, prop, tuple(w for w in words if w))
    obj.Enabled = True


def _add_recipients(obj: Any, people: list[str]) -> list[str]:
    for p in people:
        obj.Recipients.Add(p)
    unresolved = []
    if not obj.Recipients.ResolveAll():
        unresolved = [str(_safe_get(r, "Name", "")) for r in obj.Recipients if not _safe_get(r, "Resolved", True)]
    obj.Enabled = True
    return unresolved


def _apply_conditions(container: Any, c: RuleConditions) -> list[str]:
    unresolved: list[str] = []
    for field, (prop, kind, value_prop) in _CONDITIONS.items():
        value = getattr(c, field)
        if not value:
            continue
        obj = getattr(container, prop)
        if kind == "flag":
            obj.Enabled = True
        elif kind == "recipients":
            unresolved += _add_recipients(obj, value)
        elif kind == "words":
            _set_words(obj, value_prop, value)
        elif kind == "importance":
            obj.Importance = _IMPORTANCE[value]
            obj.Enabled = True
    return unresolved


def _apply_actions(namespace: Any, container: Any, a: RuleActions) -> list[str]:
    unresolved: list[str] = []
    for field, (prop, kind) in _ACTIONS.items():
        value = getattr(a, field)
        if not value:
            continue
        obj = getattr(container, prop)
        if kind == "flag":
            obj.Enabled = True
        elif kind == "folder":
            obj.Folder = resolve_folder(namespace, value)
            obj.Enabled = True
        elif kind == "categories":
            from outlook_mcp.client.categories import ensure_categories

            ensure_categories(namespace, list(value), a.category_color)
            obj.Categories = tuple(value)
            obj.Enabled = True
        elif kind == "mark":
            obj.MarkInterval = _MARK_INTERVAL[value]
            obj.Enabled = True
        elif kind == "recipients":
            unresolved += _add_recipients(obj, value)
    return unresolved


def _summary(name: str, c: RuleConditions, e: RuleConditions | None, a: RuleActions) -> dict[str, Any]:
    clean = lambda m: {k: v for k, v in m.model_dump(exclude_defaults=True).items() if k != "category_color"}  # noqa: E731
    return {
        "name": name,
        "conditions": clean(c),
        "exceptions": clean(e) if e else {},
        "actions": clean(a),
    }


def create_rule(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    conditions: RuleConditions,
    actions: RuleActions,
    exceptions: RuleConditions | None = None,
    enabled: bool = True,
    replace: bool = False,
    run_now: bool = False,
    confirm: bool = False,
) -> dict[str, Any]:
    name = name.strip()
    if not name:
        raise OutlookError("Rule name is empty.")
    if not actions.model_dump(exclude_defaults=True, exclude={"category_color"}):
        raise OutlookError("A rule needs at least one action.")
    if actions.move_to_folder and (actions.delete or actions.delete_permanently):
        raise OutlookError("Choose either move_to_folder or delete, not both.")
    _check_sending_actions(namespace, actions)
    for folder in (actions.move_to_folder, actions.copy_to_folder):
        if folder:
            resolve_folder(namespace, folder)  # fail early on a bad folder

    existing = _find_rule(_rules(namespace), name)
    if existing is not None and not replace:
        raise OutlookError(
            f"A rule named '{existing[1].Name}' already exists. Pass replace=true to replace it, "
            "or choose another name."
        )
    summary = _summary(name, conditions, exceptions, actions)
    notes = []
    if not conditions.model_dump(exclude_defaults=True):
        notes.append("No conditions: this rule applies to EVERY incoming message.")
    if actions.delete_permanently:
        notes.append("delete_permanently skips Deleted Items — matching mail can't be recovered.")

    if not confirm:
        out = {
            "status": "preview",
            "rule": summary,
            "matches": preview_matches(namespace, conditions, exceptions),
            "replaces_existing": existing is not None,
            "next": "Show this to the user. Call again with confirm=true to create the rule.",
        }
        if notes:
            out["notes"] = notes
        return out

    rules = _rules(namespace)  # fresh, right before saving
    if replace:
        again = _find_rule(rules, name)
        if again is not None:
            rules.Remove(again[0])
    rule = rules.Create(name, OL_RULE_RECEIVE)
    unresolved = _apply_conditions(rule.Conditions, conditions)
    if exceptions is not None:
        unresolved += _apply_conditions(rule.Exceptions, exceptions)
    unresolved += _apply_actions(namespace, rule.Actions, actions)
    if unresolved:
        raise OutlookError(
            f"Not saved: couldn't resolve {', '.join(unresolved)} in the address book. "
            "Use full email addresses (outlook_resolve_name can help)."
        )
    rule.Enabled = bool(enabled)
    try:
        rules.Save(False)
    except Exception as exc:
        raise OutlookError(
            f"Outlook refused to save the rule: {exc}. Some condition/action combinations "
            "aren't allowed, or the mailbox's rule quota is full."
        ) from exc
    saved = _find_rule(_rules(namespace), name)
    out = {
        "status": "replaced" if existing is not None else "created",
        "rule": describe_rule(saved[1], saved[0]) if saved else summary,
    }
    if saved and _safe_get(saved[1], "IsLocalRule", False):
        notes.append("Outlook made this a client-only rule: it runs only while Outlook is open.")
    if run_now and saved:
        out["ran_on_inbox"] = _execute(namespace, saved[1], "inbox", False, "all")
    if notes:
        out["notes"] = notes
    return out


def delete_rule(outlook: Any, namespace: Any, *, name: str, confirm: bool = False) -> dict[str, Any]:
    rules = _rules(namespace)
    found = _find_rule(rules, name)
    if found is None:
        raise OutlookError(f"Rule '{name}' not found. Use outlook_list_rules.")
    if not confirm:
        return {
            "status": "preview",
            "rule": describe_rule(found[1], found[0]),
            "next": "Call again with confirm=true to delete it.",
        }
    rules.Remove(found[0])
    rules.Save(False)
    return {"status": "deleted", "name": found[1].Name}


def _conditions_from_rule(rule: Any) -> tuple[RuleConditions, RuleConditions | None]:
    d = describe_rule(rule, 0)
    c = RuleConditions(**d["conditions"])
    e = RuleConditions(**d["exceptions"]) if d["exceptions"] else None
    return c, e


def _execute(namespace: Any, rule: Any, folder: str, include_subfolders: bool, which: str) -> dict[str, Any]:
    f = resolve_folder(namespace, folder)
    rule.Execute(False, f, bool(include_subfolders), _EXECUTE_OPTION[which])
    return {"folder": _safe_get(f, "Name"), "include_subfolders": include_subfolders, "messages": which}


def run_rule(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    folder: str = "inbox",
    include_subfolders: bool = False,
    which: str = "all",
    confirm: bool = False,
) -> dict[str, Any]:
    """Apply an existing rule to mail already in a folder ("Run Rules Now")."""
    if which not in _EXECUTE_OPTION:
        raise OutlookError("which must be one of: 'all', 'read', 'unread'.")
    found = _find_rule(_rules(namespace), name)
    if found is None:
        raise OutlookError(f"Rule '{name}' not found. Use outlook_list_rules.")
    rule = found[1]
    desc = describe_rule(rule, found[0])
    sending = [f for f in _SENDING_ACTIONS if desc["actions"].get(f)]
    if sending and draft_only():
        raise OutlookError(
            "Refused: this rule forwards/redirects mail and the server is in draft-only mode."
        )
    if not confirm:
        c, e = _conditions_from_rule(rule)
        return {
            "status": "preview",
            "rule": desc,
            "matches": preview_matches(namespace, c, e, folder=folder, days=3650, max_scan=2000),
            "next": "Running applies the actions to these messages. Call again with confirm=true.",
        }
    return {"status": "ran", "rule": desc["name"], **_execute(namespace, rule, folder, include_subfolders, which)}
