"""Strip quoted history from a reply's plain-text body.

When a whole conversation is returned, every reply repeats the messages
before it. Cutting each body at the start of its quoted part keeps a
thread readable and small. Detection is heuristic, so the full body can
always be requested instead.

Recognized starts of quoted history (at the beginning of a line):

* Outlook's reply header: a line of underscores, or "From:" (and the
  localized forms, e.g. Hebrew "מאת:") followed within a few lines by
  "Sent:" / "Date:" / "שנשלח:" / "נשלח:" etc.;
* "-----Original Message-----" (and "-----הודעה מקורית-----");
* "On <date>, <name> wrote:" (Gmail/Apple style);
* a block of lines starting with ">".
"""

from __future__ import annotations

import re

_FROM = r"(?:From|Von|De|Da|Van|Från|Fra|Od|От|מאת|من)"
_SENT = r"(?:Sent|Date|Gesendet|Envoyé|Enviado|Inviato|Verzonden|Skickat|Sendt|Отправлено|נשלח|תאריך|التاريخ|أرسل)"

_HEADER = re.compile(
    rf"^[ \t]*\*?{_FROM}\*?[ \t]*:.*\n(?:.*\n){{0,3}}?[ \t]*\*?{_SENT}\*?[ \t]*:",
    re.MULTILINE,
)
_ORIGINAL = re.compile(r"^[ \t]*-{2,}[ \t]*(?:Original Message|הודעה מקורית|Message d'origine|Ursprüngliche Nachricht)[ \t]*-{2,}", re.MULTILINE | re.IGNORECASE)
_WROTE = re.compile(r"^[ \t]*On .{4,200}(?:\n.{0,200})? wrote:[ \t]*$", re.MULTILINE)
_UNDERSCORES = re.compile(r"^[ \t]*_{10,}[ \t]*$", re.MULTILINE)
_QUOTE_BLOCK = re.compile(r"^(?:>.*\n?){2,}", re.MULTILINE)


def split_quoted(body: str) -> tuple[str, bool]:
    """Return ``(new_text, had_quote)``: the body up to its quoted history."""
    if not body:
        return "", False
    text = body.replace("\r\n", "\n").replace("\r", "\n")
    cuts = []
    for rx in (_HEADER, _ORIGINAL, _WROTE, _QUOTE_BLOCK):
        m = rx.search(text)
        if m:
            cuts.append(m.start())
    if not cuts:
        return text.strip(), False
    cut = min(cuts)
    # Outlook puts a line of underscores right above its header: cut there.
    u = None
    for m in _UNDERSCORES.finditer(text, 0, cut + 1):
        u = m
    if u is not None and len(text[u.end():cut].strip()) == 0:
        cut = u.start()
    head = text[:cut].rstrip()
    return head, True
