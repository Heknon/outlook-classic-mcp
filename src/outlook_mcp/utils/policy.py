"""Server-wide safety switches read from the environment.

**Draft-only is the default.** Unless sending is explicitly enabled, the
server cannot send anything:

* ``outlook_send_mail`` / ``outlook_reply_mail`` / ``outlook_forward_mail``
  always save to Drafts, whatever ``save_only`` says;
* ``outlook_send_draft`` is refused;
* meeting invites and updates are saved but not sent, cancellations are
  refused, and meeting responses are left in Drafts.

The user reviews and sends from Outlook.

To allow sending, set ``OUTLOOK_MCP_ALLOW_SEND=1`` (or, equivalently,
``OUTLOOK_MCP_DRAFT_ONLY=0``). ``OUTLOOK_MCP_DRAFT_ONLY=1`` always wins
and forces draft-only even if sending is allowed elsewhere.
"""

from __future__ import annotations

import os

_TRUE = ("1", "true", "yes", "on")
_FALSE = ("0", "false", "no", "off")


def _env(name: str) -> str:
    return os.environ.get(name, "").strip().lower()


def draft_only() -> bool:
    explicit = _env("OUTLOOK_MCP_DRAFT_ONLY")
    if explicit in _TRUE:
        return True
    if explicit in _FALSE:
        return False
    return _env("OUTLOOK_MCP_ALLOW_SEND") not in _TRUE


DRAFT_ONLY_NOTE = (
    "Draft-only mode is on (the default), so nothing was sent. Review and "
    "send it from Outlook. To let the server send, set OUTLOOK_MCP_ALLOW_SEND=1."
)
