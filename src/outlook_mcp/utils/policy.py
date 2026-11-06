"""Server-wide safety switches read from the environment.

``OUTLOOK_MCP_DRAFT_ONLY=1`` makes the server unable to send anything:

* ``outlook_send_mail`` / ``outlook_reply_mail`` / ``outlook_forward_mail``
  always save to Drafts, whatever ``save_only`` says;
* ``outlook_send_draft`` is refused;
* meeting invites, meeting updates and cancellations are saved but not
  sent, and meeting responses are left in Drafts.

The user reviews and sends from Outlook. This is the safest way to let
an agent work against a corporate mailbox.
"""

from __future__ import annotations

import os

_TRUE = ("1", "true", "yes", "on")


def draft_only() -> bool:
    return os.environ.get("OUTLOOK_MCP_DRAFT_ONLY", "").strip().lower() in _TRUE


DRAFT_ONLY_NOTE = (
    "Draft-only mode is on (OUTLOOK_MCP_DRAFT_ONLY), so nothing was sent. "
    "Review and send it from Outlook."
)
