"""Pure helpers for attachments: safe file names, kinds, default folder,
and links to files that live elsewhere (SharePoint, OneDrive, file shares).
"""

from __future__ import annotations

import os
import re
from pathlib import Path

_BAD_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')
_RESERVED = {"CON", "PRN", "AUX", "NUL", "CLOCK$"} | {f"COM{i}" for i in range(1, 10)} | {
    f"LPT{i}" for i in range(1, 10)
}

KINDS = {
    "pdf": "pdf",
    "doc": "word", "docx": "word", "docm": "word", "dot": "word", "dotx": "word", "rtf": "word", "odt": "word",
    "ppt": "powerpoint", "pptx": "powerpoint", "pptm": "powerpoint", "pps": "powerpoint", "ppsx": "powerpoint",
    "pot": "powerpoint", "potx": "powerpoint", "odp": "powerpoint",
    "xls": "excel", "xlsx": "excel", "xlsm": "excel", "xlsb": "excel", "csv": "excel", "ods": "excel",
    "txt": "text", "md": "text", "json": "text", "xml": "text", "html": "text", "htm": "text", "log": "text",
    "png": "image", "jpg": "image", "jpeg": "image", "gif": "image", "bmp": "image", "tif": "image",
    "tiff": "image", "webp": "image", "heic": "image", "svg": "image",
    "msg": "email", "eml": "email",
    "zip": "archive", "7z": "archive", "rar": "archive",
    "ics": "calendar", "vcf": "contact",
}

# SharePoint / OneDrive links (cloud and on-prem) and UNC file-share paths.
_LINK = re.compile(
    r"""(https?://[^\s<>"']*(?:sharepoint|onedrive|1drv\.ms|/sites/|/personal/|/:[wxpbfo]:/)[^\s<>"']*)"""
    r"""|(\\\\[A-Za-z0-9._$-]+\\[^\s<>"'|]+)""",
    re.IGNORECASE,
)


def kind_of(filename: str) -> str:
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    return KINDS.get(ext, "other")


def safe_filename(raw: str | None, fallback: str) -> str:
    """A file name that is safe to create in a folder on Windows.

    Separators, drive colons and other invalid characters become "_", so
    the result can never point outside the target folder; reserved device
    names and dot-only names get a prefix.
    """
    name = _BAD_CHARS.sub("_", raw or "").strip().strip(".").strip()
    if not name or set(name) <= {".", "_", " "}:
        name = fallback
    stem = name.split(".", 1)[0].upper()
    if stem in _RESERVED:
        name = "_" + name
    if len(name) > 150:
        base, dot, ext = name.rpartition(".")
        name = (base[: 150 - len(ext) - 1] + "." + ext) if dot and len(ext) <= 10 else name[:150]
    return name


def folder_slug(text: str, limit: int = 60) -> str:
    slug = _BAD_CHARS.sub("_", text or "").strip().strip(".")
    slug = re.sub(r"\s+", " ", slug)
    return (slug[:limit].rstrip(" ._") or "message")


def default_attachment_root() -> Path:
    base = os.environ.get("OUTLOOK_MCP_ATTACHMENTS_DIR") or os.path.join(
        os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local"),
        "outlook-mcp",
        "attachments",
    )
    return Path(base)


def linked_files(text: str | None, limit: int = 20) -> list[str]:
    """Links in a message body to files stored elsewhere (no attachment to save)."""
    out: list[str] = []
    for m in _LINK.finditer(text or ""):
        url = (m.group(1) or m.group(2)).rstrip(".,;)>]")
        if url not in out:
            out.append(url)
        if len(out) >= limit:
            break
    return out
