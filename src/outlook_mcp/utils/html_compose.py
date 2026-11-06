"""Pure HTML helpers for composing mail bodies.

Outlook's HTML mail is Word-generated: the visible content lives inside
``<body ...><div class=WordSection1> ... </div></body>`` and the user's
compose font is attached to the ``MsoNormal`` paragraph class in the
``<head>`` styles. New text therefore has to go *inside* that structure,
as ``MsoNormal`` paragraphs, or Outlook renders it with its fallback
font and spacing (the upstream bug where reply text was prepended before
the ``<html>`` tag).

Nothing here touches COM, so it is unit-testable on any platform.
"""

from __future__ import annotations

import re
import unicodedata
from html import escape

_BODY_OPEN_RE = re.compile(r"<body\b[^>]*>", re.IGNORECASE)
_WORDSECTION_OPEN_RE = re.compile(
    r"\s*<div\b[^>]*\bclass\s*=\s*[\"']?WordSection1\b[\"']?[^>]*>", re.IGNORECASE
)
_BODY_CLOSE_RE = re.compile(r"</body\s*>", re.IGNORECASE)

# Outlook wraps the signature it inserts automatically in this bookmark.
AUTO_SIGNATURE_MARKER = "_MailAutoSig"

EMPTY_PARAGRAPH = "<p class=MsoNormal>&nbsp;</p>"


def base_direction(text: str) -> str | None:
    """'rtl' or 'ltr' from the first strongly directional character.

    The same rule as HTML ``dir=auto`` / Word's paragraph detection, so a
    line like "שלום, see the attached PDF" is right-to-left and
    "Re: הצעת מחיר" is left-to-right. None when there is no letter at all.
    """
    for ch in text:
        bidi = unicodedata.bidirectional(ch)
        if bidi in ("R", "AL"):
            return "rtl"
        if bidi == "L":
            return "ltr"
    return None


def has_rtl(text: str) -> bool:
    return any(unicodedata.bidirectional(ch) in ("R", "AL") for ch in text or "")


# What Outlook/Word write for a right-to-left paragraph.
_RTL_P = "<p class=MsoNormal dir=RTL style='text-align:right;direction:rtl;unicode-bidi:embed'>"


def text_to_html(text: str) -> str:
    """Convert plain text to Outlook-style ``MsoNormal`` paragraphs.

    One paragraph per line; blank lines become empty paragraphs so the
    spacing the author typed is kept. Text is HTML-escaped. Lines whose
    first letter is Hebrew/Arabic become right-to-left paragraphs (blank
    lines between them too, so the gap sits on the right side).
    """
    if not text:
        return ""
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    dirs = [base_direction(line) for line in lines]
    out = []
    prev_dir = None
    for line, direction in zip(lines, dirs):
        if line.strip():
            # Keep runs of spaces (indentation, aligned text) visible.
            safe = escape(line, quote=False).replace("  ", " &nbsp;")
            if direction == "rtl":
                out.append(f"{_RTL_P}<span lang=HE>{safe}</span></p>")
            else:
                out.append(f"<p class=MsoNormal>{safe}</p>")
            prev_dir = direction or prev_dir
        elif prev_dir == "rtl":
            out.append(f"{_RTL_P}&nbsp;</p>")
        else:
            out.append(EMPTY_PARAGRAPH)
    return "".join(out)


def body_insert_position(html: str) -> int | None:
    """Index where new content should go at the top of the visible body.

    That is right after ``<body ...>`` and, when present, right after the
    opening ``<div class=WordSection1>`` that Word-generated mail wraps
    its content in. Returns None when the document has no ``<body>``.
    """
    m = _BODY_OPEN_RE.search(html)
    if not m:
        return None
    pos = m.end()
    w = _WORDSECTION_OPEN_RE.match(html, pos)
    if w:
        pos = w.end()
    return pos


def insert_at_top(existing_html: str | None, fragment: str) -> str:
    """Insert ``fragment`` at the top of the visible body of ``existing_html``."""
    if not fragment:
        return existing_html or ""
    if not existing_html or not existing_html.strip():
        return f"<html><body>{fragment}</body></html>"
    pos = body_insert_position(existing_html)
    if pos is None:
        # Not a full document (a bare fragment): just prepend.
        return fragment + existing_html
    return existing_html[:pos] + fragment + existing_html[pos:]


def append_to_body(existing_html: str, fragment: str) -> str:
    """Insert ``fragment`` at the end of the visible body (before ``</body>``)."""
    m = _BODY_CLOSE_RE.search(existing_html)
    if not m:
        return existing_html + fragment
    return existing_html[: m.start()] + fragment + existing_html[m.start() :]


# Used only for documents we build from scratch (new mail with a signature).
# Outlook's own messages carry the user's compose styles; ours need at least
# tight paragraph spacing so MsoNormal paragraphs don't get double gaps.
_NEW_DOC_STYLE = (
    "<style>p.MsoNormal{margin:0cm;font-family:Aptos,Calibri,Arial,sans-serif;"
    "font-size:11.0pt}</style>"
)


def new_html_document(fragment: str) -> str:
    return (
        "<html><head><meta http-equiv=Content-Type content=\"text/html; charset=utf-8\">"
        f"{_NEW_DOC_STYLE}</head><body><div class=WordSection1>{fragment}</div></body></html>"
    )


def body_inner_html(html: str) -> str:
    """Return the markup between ``<body ...>`` and ``</body>``.

    Falls back to the whole string when there is no body element.
    """
    m = _BODY_OPEN_RE.search(html)
    if not m:
        return html
    end = _BODY_CLOSE_RE.search(html, m.end())
    return html[m.end() : end.start() if end else len(html)]


def has_auto_signature(html: str | None) -> bool:
    """True when Outlook already inserted a signature into this item."""
    return bool(html) and AUTO_SIGNATURE_MARKER in html


def compose_fragment(
    body: str,
    *,
    is_html: bool,
    signature_html: str | None = None,
    trailing_gap: bool = False,
) -> str:
    """Build the block to insert: the body, then (optionally) a signature.

    ``trailing_gap`` adds an empty paragraph after the block, which keeps
    a visual gap before quoted history in replies and forwards.
    """
    parts: list[str] = []
    if body:
        parts.append(body if is_html else text_to_html(body))
    if signature_html:
        if parts:
            parts.append(EMPTY_PARAGRAPH)
        parts.append(signature_html)
    if parts and trailing_gap:
        parts.append(EMPTY_PARAGRAPH)
    return "".join(parts)
