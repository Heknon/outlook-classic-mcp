"""Export your Outlook Quick Steps (read-only) so their format can be checked.

Outlook stores each Quick Step as a hidden item in a hidden configuration
folder of your mailbox ("Quick Step Settings"), with the definition as XML
in a MAPI property. That format isn't officially documented, so this
script finds it without assuming too much, and writes what it finds to
%USERPROFILE%\\Documents\\outlook-quick-steps.json.

Run on the Windows PC with classic Outlook open:

    python scripts\\export_quick_steps.py

It only reads: nothing in your mailbox is changed. The output contains
your Quick Step names, their XML (folder IDs, category names and any
addresses the steps use) and the names of the folders they point to. No
email content. Review it before sharing.
"""

from __future__ import annotations

import binascii
import json
import os
import re
import sys
from pathlib import Path

if sys.platform != "win32":
    sys.exit("Run this on Windows with classic Outlook.")

import win32com.client  # noqa: E402

PROP = "http://schemas.microsoft.com/mapi/proptag/"
CONTAINER_CLASS = PROP + "0x3613001F"  # PR_CONTAINER_CLASS
CANDIDATE_PROPS = {
    "PR_ROAMING_XMLSTREAM": PROP + "0x7C080102",
    "PR_ROAMING_DICTIONARY": PROP + "0x7C070102",
    "PR_ROAMING_DATATYPES": PROP + "0x7C060003",
}
OL_HIDDEN_ITEMS = 1  # OlTableContents.olHiddenItems
HEX_ID = re.compile(r"\b[0-9A-Fa-f]{40,}\b")


def get(obj, name, default=None):
    try:
        return getattr(obj, name)
    except Exception:  # noqa: BLE001
        return default


def prop(obj, tag):
    try:
        return obj.PropertyAccessor.GetProperty(tag)
    except Exception as exc:  # noqa: BLE001
        return exc


def decode(value):
    """Binary property -> text (UTF-8 or UTF-16), else hex."""
    if isinstance(value, Exception) or value is None:
        return None, None
    raw = bytes(value) if not isinstance(value, (bytes, bytearray)) else bytes(value)
    for enc in ("utf-8-sig", "utf-16"):
        try:
            text = raw.decode(enc)
            if "<" in text and ">" in text:
                return text, enc
        except UnicodeDecodeError:
            continue
    return binascii.hexlify(raw[:4096]).decode(), "hex"


def candidate_folders(root):
    """Hidden configuration folders directly under the mailbox root."""
    out = []
    for f in list(root.Folders):
        cls = prop(f, CONTAINER_CLASS)
        cls = None if isinstance(cls, Exception) else cls
        name = str(get(f, "Name", ""))
        if cls == "IPF.Configuration" or "quick" in name.lower():
            out.append((f, name, cls))
    return out


def hidden_items(ns, folder, store_id):
    rows = []
    try:
        table = folder.GetTable("", OL_HIDDEN_ITEMS)
    except Exception as exc:  # noqa: BLE001
        return [{"error": f"GetTable(hidden) failed: {exc}"}]
    while not table.EndOfTable:
        row = table.GetNextRow()
        eid = row.Item("EntryID")
        entry = {"entry_id": eid, "subject": row.Item("Subject"), "message_class": row.Item("MessageClass")}
        try:
            item = ns.GetItemFromID(eid, store_id)
        except Exception as exc:  # noqa: BLE001
            entry["open_error"] = str(exc)
            rows.append(entry)
            continue
        for label, tag in CANDIDATE_PROPS.items():
            value = prop(item, tag)
            if label == "PR_ROAMING_DATATYPES":
                entry[label] = None if isinstance(value, Exception) else value
                continue
            text, enc = decode(value)
            if text is not None:
                entry[label] = {"encoding": enc, "content": text}
        rows.append(entry)
    return rows


def resolve_ids(ns, text, store_id):
    found = {}
    for hex_id in set(HEX_ID.findall(text or "")):
        for args in ((hex_id, store_id), (hex_id,)):
            try:
                f = ns.GetFolderFromID(*args)
                found[hex_id] = str(get(f, "FolderPath") or get(f, "Name"))
                break
            except Exception:  # noqa: BLE001
                continue
    return found


def main() -> int:
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    store = ns.DefaultStore
    store_id = store.StoreID
    root = store.GetRootFolder()
    report = {
        "outlook_version": str(get(outlook, "Version")),
        "store": str(get(store, "DisplayName")),
        "root_children": [str(get(f, "Name")) for f in list(root.Folders)],
        "folders": [],
    }
    for folder, name, cls in candidate_folders(root):
        items = hidden_items(ns, folder, store_id)
        for it in items:
            xml = (it.get("PR_ROAMING_XMLSTREAM") or {}).get("content")
            if xml:
                it["resolved_folder_ids"] = resolve_ids(ns, xml, store_id)
        report["folders"].append({"name": name, "container_class": cls, "hidden_items": items})

    out = Path(os.environ.get("USERPROFILE", ".")) / "Documents" / "outlook-quick-steps.json"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")

    n_xml = sum(1 for f in report["folders"] for i in f["hidden_items"] if "PR_ROAMING_XMLSTREAM" in i)
    print(f"Checked {len(report['folders'])} configuration folder(s); found {n_xml} item(s) with XML.")
    for f in report["folders"]:
        print(f"  {f['name']} ({f['container_class']}): {len(f['hidden_items'])} hidden item(s)")
    print(f"Saved to {out}")
    if not n_xml:
        print("No Quick Step XML found — the file still helps: it lists what the mailbox root contains.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
