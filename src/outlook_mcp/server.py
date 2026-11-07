"""FastMCP server construction and tool registration."""

from __future__ import annotations

from mcp.server.fastmcp import FastMCP

from outlook_mcp.bridge import OutlookBridge
from outlook_mcp.tools import register_all
from outlook_mcp.ui import register_ui

INSTRUCTIONS = """\
This MCP server gives you full access to Microsoft Outlook desktop on
Windows via COM automation. It can read mail and write drafts (and send,
if the user enabled it), manage calendar events, check colleagues'
availability, look up people, and work with contacts, tasks, color
categories, mail rules and Out-of-Office status.

DRAFT-ONLY BY DEFAULT: unless the user set OUTLOOK_MCP_ALLOW_SEND=1,
nothing can be sent. Mail, replies and forwards are saved to Drafts
(results say status "saved_to_drafts"), meeting invites and updates are
saved unsent, and send_draft is refused. Tell the user where the draft
is; don't try to work around it.

Authentication piggybacks on whatever account Outlook is signed into —
no Microsoft Graph API, no Entra app registration, no OAuth tokens.

PREREQUISITE: classic Outlook (OUTLOOK.EXE) must be installed. The
server auto-launches Outlook on first call; the new "modern" Outlook
(olk.exe) is NOT supported.

Tool categories (all prefixed `outlook_`):
  - Mail: list, search, search_mails_by_person, get, get_thread (a whole
    conversation in one call, quoted history removed), send, reply, forward
    (all three take save_only=true for drafts), send_draft, list_signatures,
    move, delete, mark, save_attachments (saves files and returns their
    paths; read them with your own file tools)
  - Inbox zero: list_quick_steps (the user's Quick Steps = their filing
    buckets), run_quick_step, move_mails (bulk), check_waiting (which
    Waiting items got replies). Propose filing to the user before moving.
  - Folders: list_folders, create_folder, list_mailboxes (yours, online
    archive, shared mailboxes, PSTs). Mail searches take search_in=
    'mailbox' | 'archive' | 'all' or mailboxes=[...] to search every folder
    of those mailboxes, not just one folder.
  - Calendar: list_events (also a colleague's shared calendar via mailbox=),
    get_event, create_event, update_event, delete_event, respond_event,
    get_free_busy (anyone in the org), find_meeting_times (scheduling assistant)
  - Contacts: list_contacts, search_contacts (incl. org directory/GAL),
    get_contact, resolve_name (like Ctrl+K: name -> address, or candidates
    to ask the user about when ambiguous)
  - Tasks: list_tasks, create_task, complete_task
  - Categories (colour labels on mail, meetings, tasks): list_categories,
    create_category, update_category, delete_category, set_category
    (replace/add/remove, bulk, one meeting occurrence; creates missing ones)
  - Rules: list_rules (details=true explains them), create_rule, run_rule,
    delete_rule, toggle_rule. create/run/delete return a PREVIEW unless
    confirm=true: show it to the user first. Forward/redirect rules are
    refused in draft-only mode and limited to the user's own domain.
  - Out-of-Office: get_out_of_office
  - Account: whoami  (sanity check on the bound mailbox)

Most read tools accept response_format='markdown' (default) or 'json'.
Item references use Outlook EntryID strings; list tools return them on
every item — pass them back to detail/edit/delete tools. For one
occurrence of a recurring meeting pass occurrence_start (its listed
start); changing a whole series needs apply_to_series=true.

All datetimes are in the USER'S LOCAL timezone with an explicit UTC
offset (e.g. 2026-06-10T16:33:22+05:00). Present them as-is — do NOT
convert to another timezone. outlook_whoami reports the timezone name
and current local time.
"""


def build_server() -> tuple[FastMCP, OutlookBridge]:
    """Construct the FastMCP instance, bridge, and wire all tools.

    The bridge is *not* started here — entrypoint.main() does that so
    the readiness wait happens after logging is configured.
    """
    mcp = FastMCP("outlook_mcp", instructions=INSTRUCTIONS)
    bridge = OutlookBridge()
    register_all(mcp, bridge)
    register_ui(mcp)
    return mcp, bridge
