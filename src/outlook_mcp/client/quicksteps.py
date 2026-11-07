"""Read Outlook Quick Steps and carry out their (safe) actions.

There is no Outlook API for Quick Steps. Outlook keeps each one as a
hidden item in a hidden configuration folder of the mailbox ("Quick Step
Settings", folder class IPF.Configuration) with its definition as XML in
PR_ROAMING_XMLSTREAM. That format isn't documented, so parsing is
deliberately conservative:

* action elements are recognized by name (the kinds Outlook supports:
  MoveToFolder, CopyToFolder, MarkAsRead, CategorizeMessage, FlagMessage,
  Forward, Reply, NewMeeting, ...);
* folder targets are recognized as hex EntryIDs and resolved with
  GetFolderFromID; categories by name or by category GUID;
* a Quick Step is runnable only when every action was understood and is
  one of the safe kinds below. Anything else is shown, never guessed.

Runnable: move, copy, mark read/unread, set importance, categorize,
clear categories, flag, clear flag, mark complete; delete with
confirm=true. Not run: reply/forward/new message/meeting/task/
appointment steps and conversation-wide "always move"/"ignore".
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.errors import OutlookError

_PROP = "http://schemas.microsoft.com/mapi/proptag/"
CONTAINER_CLASS = _PROP + "0x3613001F"
XML_STREAM = _PROP + "0x7C080102"
OL_HIDDEN_ITEMS = 1

_HEX_ID = re.compile(r"^[0-9A-Fa-f]{40,}$")
_GUID = re.compile(r"^\{?[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}?$")

# Normalized element name -> action kind.
_KINDS = {
    "movetofolder": "move",
    "copytofolder": "copy",
    "deletemessage": "delete",
    "delete": "delete",
    "permanentlydeletemessage": "delete_permanently",
    "markasread": "mark_read",
    "markasunread": "mark_unread",
    "setimportance": "importance",
    "categorizemessage": "categorize",
    "categorize": "categorize",
    "clearcategories": "clear_categories",
    "flagmessage": "flag",
    "flag": "flag",
    "clearflag": "clear_flag",
    "markcomplete": "mark_complete",
    # Recognized but not run by this server:
    "newmessage": "new_message",
    "forward": "forward",
    "forwardmessageasattachment": "forward_as_attachment",
    "reply": "reply",
    "replyall": "reply_all",
    "replywithmeeting": "reply_with_meeting",
    "newmeeting": "new_meeting",
    "createtaskwithattachment": "create_task",
    "createtaskwithtext": "create_task",
    "createappointmentwithattachment": "create_appointment",
    "createappointmentwithtext": "create_appointment",
    "alwaysmovemessagesinthisconversation": "always_move_conversation",
    "ignoremessagesinthisconversation": "ignore_conversation",
}
SAFE_KINDS = {"move", "copy", "mark_read", "mark_unread", "importance", "categorize",
              "clear_categories", "flag", "clear_flag", "mark_complete"}
CONFIRM_KINDS = {"delete", "delete_permanently"}
_NEEDS_FOLDER = {"move", "copy", "always_move_conversation"}


def _norm(tag: str) -> str:
    tag = tag.rsplit("}", 1)[-1].lower()
    for prefix in ("qak", "action"):
        if tag.startswith(prefix) and tag[len(prefix):] in _KINDS:
            return tag[len(prefix):]
    return tag


def _texts(el: ET.Element) -> list[str]:
    out = []
    for node in el.iter():
        if node.text and node.text.strip():
            out.append(node.text.strip())
        out.extend(str(v).strip() for v in node.attrib.values() if str(v).strip())
    return out


def _named(el: ET.Element, *words: str) -> list[str]:
    """Text/attribute values of descendants whose name contains a word."""
    out = []
    for node in el.iter():
        name = node.tag.rsplit("}", 1)[-1].lower()
        if any(w in name for w in words) and node.text and node.text.strip():
            out.append(node.text.strip())
        for k, v in node.attrib.items():
            if any(w in k.lower() for w in words) and str(v).strip():
                out.append(str(v).strip())
    return out


def parse_quick_step_xml(xml: str) -> dict[str, Any]:
    """Pure parse: name and a list of raw actions (kind + parameters)."""
    root = ET.fromstring(xml)
    name = root.attrib.get("Name") or root.attrib.get("name")
    if not name:
        found = _named(root, "name")
        name = found[0] if found else None
    actions = []
    for el in root.iter():
        key = _norm(el.tag)
        if el is root or key not in _KINDS:
            continue
        values = _texts(el)
        action: dict[str, Any] = {"kind": _KINDS[key], "element": el.tag.rsplit("}", 1)[-1]}
        ids = [v for v in values if _HEX_ID.match(v)]
        if ids:
            action["folder_ids"] = ids
        cats = _named(el, "categor")
        if cats:
            action["category_refs"] = cats
        imp = _named(el, "importance", "level")
        if imp:
            action["importance_raw"] = imp[0]
        flag = _named(el, "flag", "interval", "mark")
        if flag:
            action["flag_raw"] = flag[0]
        actions.append(action)
    return {"name": name, "actions": actions}


# ------------------------------------------------------------- mailbox


def _quick_step_folder(namespace: Any) -> Any:
    root = namespace.DefaultStore.GetRootFolder()
    for f in list(root.Folders):
        try:
            cls = f.PropertyAccessor.GetProperty(CONTAINER_CLASS)
        except Exception:
            cls = None
        name = str(_safe_get(f, "Name", "")).lower()
        if "quick step" in name or (cls == "IPF.Configuration" and "quick" in name):
            return f
    raise OutlookError(
        "Couldn't find the Quick Steps settings folder in your mailbox. Run "
        "scripts/export_quick_steps.py and share the result so support can be added."
    )


def _decode(value: Any) -> str | None:
    if value is None:
        return None
    raw = bytes(value)
    for enc in ("utf-8-sig", "utf-16"):
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        if "<" in text:
            return text
    return None


def _raw_quick_steps(namespace: Any) -> list[tuple[str, str]]:
    folder = _quick_step_folder(namespace)
    store_id = _safe_get(namespace.DefaultStore, "StoreID")
    table = folder.GetTable("", OL_HIDDEN_ITEMS)
    out = []
    while not table.EndOfTable:
        row = table.GetNextRow()
        eid = row.Item("EntryID")
        try:
            item = namespace.GetItemFromID(eid, store_id)
            xml = _decode(item.PropertyAccessor.GetProperty(XML_STREAM))
        except Exception:
            continue
        if xml:
            out.append((str(row.Item("Subject") or ""), xml))
    return out


def _category_name(namespace: Any, ref: str) -> str:
    if _GUID.match(ref):
        wanted = ref.strip("{}").lower()
        cats = namespace.Categories
        for i in range(cats.Count):
            c = cats.Item(i + 1)
            if str(_safe_get(c, "CategoryID", "")).strip("{}").lower() == wanted:
                return str(c.Name)
    return ref


def _resolve(namespace: Any, raw: dict[str, Any], fallback_name: str) -> dict[str, Any]:
    store_id = _safe_get(namespace.DefaultStore, "StoreID")
    actions, problems = [], []
    for a in raw["actions"]:
        act: dict[str, Any] = {"kind": a["kind"]}
        if a["kind"] in _NEEDS_FOLDER:
            folder = None
            for fid in a.get("folder_ids", []):
                for args in ((fid, store_id), (fid,)):
                    try:
                        folder = namespace.GetFolderFromID(*args)
                        break
                    except Exception:
                        continue
                if folder is not None:
                    break
            if folder is None:
                problems.append(f"{a['element']}: target folder not understood")
            else:
                act["folder"] = str(_safe_get(folder, "FolderPath") or _safe_get(folder, "Name"))
                act["_folder"] = folder
        if a["kind"] == "categorize":
            names = [_category_name(namespace, r) for r in a.get("category_refs", [])]
            names = [n for n in names if n and not _HEX_ID.match(n)]
            if not names:
                problems.append(f"{a['element']}: categories not understood")
            act["categories"] = names
        if a["kind"] == "importance":
            raw_imp = str(a.get("importance_raw", "")).lower()
            imp = {"0": 0, "low": 0, "1": 1, "normal": 1, "2": 2, "high": 2}.get(raw_imp)
            if imp is None:
                problems.append(f"{a['element']}: importance not understood")
            act["importance"] = imp
        if a["kind"] not in SAFE_KINDS | CONFIRM_KINDS:
            problems.append(f"{a['element']}: '{a['kind']}' steps aren't run by this server")
        actions.append(act)
    if not actions:
        problems.append("no actions recognized")
    public = [{k: v for k, v in act.items() if not k.startswith("_")} for act in actions]
    return {
        "name": raw.get("name") or fallback_name,
        "actions": public,
        "runnable": not problems,
        "needs_confirm": any(a["kind"] in CONFIRM_KINDS for a in actions),
        "problems": problems,
        "_actions": actions,
    }


def load_quick_steps(namespace: Any) -> list[dict[str, Any]]:
    steps = []
    for subject, xml in _raw_quick_steps(namespace):
        try:
            raw = parse_quick_step_xml(xml)
        except ET.ParseError as exc:
            steps.append({"name": subject, "actions": [], "runnable": False,
                          "problems": [f"XML not readable: {exc}"], "_actions": []})
            continue
        steps.append(_resolve(namespace, raw, subject))
    return steps


def _public(step: dict[str, Any]) -> dict[str, Any]:
    return {k: v for k, v in step.items() if not k.startswith("_")}


def list_quick_steps(outlook: Any, namespace: Any) -> dict[str, Any]:
    steps = load_quick_steps(namespace)
    return {"count": len(steps), "quick_steps": [_public(s) for s in steps]}


def find_quick_step(namespace: Any, name: str) -> dict[str, Any]:
    steps = load_quick_steps(namespace)
    wanted = name.strip().lower()
    for s in steps:
        if str(s["name"] or "").strip().lower() == wanted:
            return s
    matches = [s for s in steps if wanted in str(s["name"] or "").lower()]
    if len(matches) == 1:
        return matches[0]
    names = ", ".join(str(s["name"]) for s in steps) or "none"
    raise OutlookError(f"Quick Step '{name}' not found. Your Quick Steps: {names}")


def _apply(item: Any, actions: list[dict[str, Any]]) -> Any:
    """Apply actions in order; returns the item (moved items change identity)."""
    from outlook_mcp.client.categories import _split

    for a in actions:
        kind = a["kind"]
        if kind == "mark_read":
            item.UnRead = False
        elif kind == "mark_unread":
            item.UnRead = True
        elif kind == "importance":
            item.Importance = a["importance"]
        elif kind == "categorize":
            cur = _split(_safe_get(item, "Categories", ""))
            item.Categories = ", ".join(cur + [c for c in a["categories"] if c not in cur])
        elif kind == "clear_categories":
            item.Categories = ""
        elif kind == "flag":
            item.FlagStatus = 2
        elif kind == "clear_flag":
            item.FlagStatus = 0
        elif kind == "mark_complete":
            item.FlagStatus = 1
        if kind in {"mark_read", "mark_unread", "importance", "categorize", "clear_categories",
                    "flag", "clear_flag", "mark_complete"}:
            item.Save()
        elif kind == "copy":
            copy = item.Copy()
            copy.Move(a["_folder"])
        elif kind == "move":
            item = item.Move(a["_folder"])
        elif kind in ("delete", "delete_permanently"):
            item.Delete()
            return None
    return item


def run_quick_step(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    entry_ids: list[str],
    confirm: bool = False,
) -> dict[str, Any]:
    step = find_quick_step(namespace, name)
    if not step["runnable"]:
        raise OutlookError(
            f"Quick Step '{step['name']}' can't be run by the server: {'; '.join(step['problems'])}. "
            "Use it in Outlook, or ask for the equivalent with outlook_move_mails / outlook_set_category."
        )
    if not entry_ids:
        raise OutlookError("Pass the entry_ids of the messages.")
    if step["needs_confirm"] and not confirm:
        return {
            "status": "preview",
            "quick_step": _public(step),
            "count": len(entry_ids),
            "next": "This Quick Step deletes messages. Call again with confirm=true once the user agrees.",
        }
    done, errors = [], []
    for eid in entry_ids:
        try:
            item = get_item_by_id(namespace, eid)
            subject = _safe_get(item, "Subject", "")
            result = _apply(item, step["_actions"])
            done.append({
                "entry_id": eid,
                "subject": subject,
                "new_entry_id": _safe_get(result, "EntryID") if result is not None else None,
            })
        except Exception as exc:  # noqa: BLE001 - report per message
            errors.append({"entry_id": eid, "error": str(exc)})
    out = {"status": "done", "quick_step": step["name"], "actions": _public(step)["actions"],
           "count": len(done), "items": done}
    if errors:
        out["errors"] = errors
    return out


def quick_step_target(namespace: Any, *words: str) -> tuple[str, Any] | None:
    """Folder that a Quick Step whose name contains one of ``words`` moves to."""
    try:
        steps = load_quick_steps(namespace)
    except OutlookError:
        return None
    for s in steps:
        if any(w in str(s["name"] or "").lower() for w in words):
            for a in s["_actions"]:
                if a["kind"] == "move" and a.get("_folder") is not None:
                    return a["folder"], a["_folder"]
    return None
