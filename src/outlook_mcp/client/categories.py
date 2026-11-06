"""Category COM operations: the colour-coded labels on mail and meetings.

Outlook categories are profile-wide (``namespace.Categories``, the
"master category list"). An item's ``Categories`` property is a string of
names joined by the Windows list separator; names on an item that aren't
in the master list show without colour, so tools here create missing
categories (with a colour) before assigning them.
"""

from __future__ import annotations

from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.errors import OutlookError

# OlCategoryColor. Names follow the enum; aliases are the labels Outlook's
# colour picker shows in recent versions.
_COLOR_CODES = {
    "none": 0,
    "red": 1,
    "orange": 2,
    "peach": 3,
    "yellow": 4,
    "green": 5,
    "teal": 6,
    "olive": 7,
    "blue": 8,
    "purple": 9,
    "maroon": 10,
    "steel": 11,
    "dark steel": 12,
    "gray": 13,
    "dark gray": 14,
    "black": 15,
    "dark red": 16,
    "dark orange": 17,
    "dark peach": 18,
    "dark yellow": 19,
    "dark green": 20,
    "dark teal": 21,
    "dark olive": 22,
    "dark blue": 23,
    "dark purple": 24,
    "dark maroon": 25,
}
_ALIASES = {
    "grey": "gray",
    "dark grey": "dark gray",
    "cranberry": "maroon",
    "brown": "dark peach",
    "gold": "dark yellow",
    "navy": "dark blue",
    "plum": "dark purple",
    "pink": "maroon",
    "light blue": "blue",
    "turquoise": "teal",
}
COLOR_NAMES = {code: name for name, code in _COLOR_CODES.items()}
COLOR_CHOICES = ", ".join(n for n in _COLOR_CODES if n != "none")


def color_code(color: str | int | None) -> int | None:
    if color is None:
        return None
    if isinstance(color, int):
        if color in COLOR_NAMES:
            return color
        raise OutlookError(f"Unknown colour number {color}; use 0-25.")
    key = " ".join(str(color).strip().lower().replace("_", " ").replace("-", " ").split())
    key = _ALIASES.get(key, key)
    if key not in _COLOR_CODES:
        raise OutlookError(f"Unknown colour '{color}'. Choose one of: {COLOR_CHOICES}.")
    return _COLOR_CODES[key]


def _all(namespace: Any) -> list[Any]:
    cats = namespace.Categories
    return [cats.Item(i + 1) for i in range(cats.Count)]


def _find(namespace: Any, name: str) -> Any | None:
    wanted = name.strip().lower()
    for cat in _all(namespace):
        if str(_safe_get(cat, "Name", "")).lower() == wanted:
            return cat
    return None


def _describe(cat: Any) -> dict[str, Any]:
    code = _safe_get(cat, "Color", 0)
    return {"name": _safe_get(cat, "Name"), "color": COLOR_NAMES.get(code, str(code)), "color_code": code}


def list_categories(outlook: Any, namespace: Any) -> dict[str, Any]:
    items = [_describe(c) for c in _all(namespace)]
    return {"count": len(items), "items": items, "colors": COLOR_CHOICES}


def create_category(outlook: Any, namespace: Any, *, name: str, color: str | int = "blue") -> dict[str, Any]:
    name = name.strip()
    if not name:
        raise OutlookError("Category name is empty.")
    if "," in name or ";" in name:
        raise OutlookError("Category names can't contain ',' or ';' (Outlook uses them as separators).")
    if _find(namespace, name) is not None:
        raise OutlookError(f"Category '{name}' already exists; use outlook_update_category to change it.")
    cat = namespace.Categories.Add(name, color_code(color))
    return {"status": "created", **_describe(cat)}


def update_category(
    outlook: Any,
    namespace: Any,
    *,
    name: str,
    new_name: str | None = None,
    color: str | int | None = None,
) -> dict[str, Any]:
    cat = _find(namespace, name)
    if cat is None:
        raise OutlookError(f"Category '{name}' not found. Use outlook_list_categories.")
    note = None
    if color is not None:
        cat.Color = color_code(color)
    if new_name:
        if _find(namespace, new_name) is not None and new_name.strip().lower() != name.strip().lower():
            raise OutlookError(f"Category '{new_name}' already exists.")
        cat.Name = new_name.strip()
        note = (
            "Renaming changes the master list only: items already tagged keep the "
            "old name until re-tagged."
        )
    out = {"status": "updated", **_describe(cat)}
    if note:
        out["note"] = note
    return out


def delete_category(outlook: Any, namespace: Any, *, name: str) -> dict[str, Any]:
    cat = _find(namespace, name)
    if cat is None:
        raise OutlookError(f"Category '{name}' not found. Use outlook_list_categories.")
    namespace.Categories.Remove(_safe_get(cat, "CategoryID") or _safe_get(cat, "Name"))
    return {
        "status": "deleted",
        "name": name,
        "note": "Items keep the label as text (shown without colour) until re-tagged.",
    }


def _split(value: str | None) -> list[str]:
    if not value:
        return []
    parts = value.replace(";", ",").split(",")
    return [p.strip() for p in parts if p.strip()]


def _target(namespace: Any, entry_id: str, occurrence_start: str | None) -> Any:
    if occurrence_start:
        from outlook_mcp.client.calendar import _target_event

        return _target_event(namespace, entry_id, occurrence_start)
    return get_item_by_id(namespace, entry_id)


def ensure_categories(namespace: Any, names: list[str], color: str | int | None) -> list[str]:
    """Create any of ``names`` missing from the master list. Returns the created ones."""
    created = []
    code = color_code(color) if color is not None else _COLOR_CODES["blue"]
    for n in names:
        if _find(namespace, n) is None:
            namespace.Categories.Add(n, code)
            created.append(n)
    return created


def set_category(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str | None = None,
    categories: str | list[str] | None = None,
    entry_ids: list[str] | None = None,
    mode: str = "replace",
    occurrence_start: str | None = None,
    create_missing: bool = True,
    color: str | int | None = None,
) -> dict[str, Any]:
    """Set, add or remove categories on one or more mails/meetings/tasks."""
    if mode not in ("replace", "add", "remove"):
        raise OutlookError("mode must be one of: 'replace', 'add', 'remove'.")
    ids = list(entry_ids or [])
    if entry_id:
        ids.insert(0, entry_id)
    if not ids:
        raise OutlookError("Pass entry_id or entry_ids.")
    if occurrence_start and len(ids) > 1:
        raise OutlookError("occurrence_start works with a single entry_id.")
    names = categories if isinstance(categories, list) else _split(categories)
    names = [n.strip() for n in names if n and n.strip()]
    if mode != "replace" and not names:
        raise OutlookError("Pass the categories to add or remove.")

    created = ensure_categories(namespace, names, color) if (create_missing and mode != "remove") else []
    results, errors = [], []
    for eid in ids:
        try:
            item = _target(namespace, eid, occurrence_start)
            current = _split(_safe_get(item, "Categories", ""))
            if mode == "replace":
                new = names
            elif mode == "add":
                new = current + [n for n in names if n.lower() not in {c.lower() for c in current}]
            else:
                drop = {n.lower() for n in names}
                new = [c for c in current if c.lower() not in drop]
            item.Categories = ", ".join(new)
            item.Save()
            results.append({"entry_id": eid, "subject": _safe_get(item, "Subject", ""), "categories": new})
        except Exception as exc:  # noqa: BLE001 - report per item in bulk
            if len(ids) == 1:
                raise
            errors.append({"entry_id": eid, "error": str(exc)})
    out: dict[str, Any] = {"status": "updated", "count": len(results), "items": results}
    if len(results) == 1 and not errors:
        # Backwards-compatible single-item shape.
        out.update({"entry_id": results[0]["entry_id"], "subject": results[0]["subject"],
                    "categories": ", ".join(results[0]["categories"])})
    if created:
        out["created_categories"] = created
    if errors:
        out["errors"] = errors
    return out
