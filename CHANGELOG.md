# Changelog

## 0.5.0

- Meeting titles where you have access: `outlook_get_free_busy` and
  `outlook_find_meeting_times` include colleagues' meeting titles/locations
  when their calendar is shared with you (`access: "calendar"`), busy/free
  otherwise; others' private appointments are masked (also in
  `outlook_list_events(mailbox=...)`).
- Attachments: `save_attachments` saves to a per-message folder by default,
  skips inline images (signature logos), saves attached emails as `.msg`,
  fixes invalid file names instead of failing, reports per-attachment
  failures, and returns path/kind/size per file. `get_mail` gives each
  attachment's kind/type/inline flag and lists SharePoint/OneDrive/file-share
  links in the body (`linked_files`).
- Inbox zero: `outlook_list_quick_steps` / `outlook_run_quick_step` read the
  user's Quick Steps from the mailbox (hidden configuration folder, XML)
  and carry out safe actions; `outlook_move_mails` (bulk);
  `outlook_check_waiting` (Waiting items that got replies). Plugin prompts
  `/outlook:triage` and `/outlook:brief`. `scripts/export_quick_steps.py`
  exports Quick Step definitions (read-only) for format checks.
- Categories: `outlook_create_category`, `outlook_update_category`
  (recolour/rename), `outlook_delete_category`; all 25 Outlook colours by
  name. `outlook_set_category` gains bulk `entry_ids`, `mode`
  (replace/add/remove), `occurrence_start` for one meeting occurrence, and
  creates missing categories in a chosen colour. `create_event` /
  `update_event` take `categories`.
- Rules: `outlook_create_rule` (conditions, exceptions, actions as typed
  objects), `outlook_run_rule` ("Run Rules Now"), `outlook_delete_rule`, all
  preview-first with `confirm=true`; `outlook_list_rules(details=true)`
  explains rules. Forward/redirect rules refused in draft-only mode and
  limited to the user's domain unless `OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD=1`.
- New `outlook_get_thread`: a whole conversation in one call, oldest first,
  across folders (incl. Sent Items and drafts), with quoted history removed
  (English and Hebrew reply headers and other common formats).
- Hebrew/Arabic text in drafts is written as right-to-left paragraphs
  (`dir=RTL`, first-strong-character rule); Hebrew new mail and Hebrew
  replies to plain-text mail use HTML. Switching a plain-text reply to HTML
  keeps the quoted history even if Outlook's conversion comes back empty.
- **Plugin installs everything itself.** `/plugin marketplace add
  Heknon/outlook-classic-mcp` + `/plugin install outlook@outlook-classic-mcp`
  is enough: `scripts/launch-mcp.cmd` installs uv per user if missing (or
  falls back to an installed Python), builds the environment in the
  plugin's persistent data folder, and starts the server from the plugin's
  own files (no git, no PyPI package). A SessionStart hook prepares the
  environment so the server starts within Claude Code's MCP timeout.
  Windows CI installs the plugin with the Claude Code CLI and connects to it.
- Search across mailboxes: `search_in` = `mailbox` / `archive` / `all`, or
  `mailboxes=[...]`, on `outlook_search_mails` and
  `outlook_search_mails_by_person`; every mail folder is searched, results
  merged newest first with folder paths; `include_deleted`. New
  `outlook_list_mailboxes` (primary, archive, shared/delegate, PST).
- End-to-end test: the real server over stdio against a simulated Outlook
  (`tests/e2e`), run in CI.
- `outlook_resolve_name` works like Ctrl+K: partial names and aliases,
  and a list of candidates (title, department, office) when several people
  match instead of just failing.
- New `outlook_search_mails_by_person`: mails from and/or to a person by
  name, across Inbox and Sent Items; returns candidates when the name is
  ambiguous.
- **Draft-only is now the default.** Nothing is sent unless
  `OUTLOOK_MCP_ALLOW_SEND=1` is set; `OUTLOOK_MCP_DRAFT_ONLY=1` still forces
  draft-only.
- `outlook_list_events` / `outlook_get_event` take `mailbox` to read a
  colleague's shared calendar.
- New `outlook_get_free_busy`: busy/tentative/out-of-office blocks for anyone
  in the organization, from Exchange free/busy (no calendar permission needed).
- New `outlook_find_meeting_times`: scheduling assistant — slots when you and
  all required attendees are free, within working hours and working days
  (configurable, e.g. Sunday–Thursday), with optional attendees reported per
  slot.

## 0.4.0 — fork of outlook-classic-mcp 0.3.2

Forked from [anasahmed07/Outlook-Classic-MCP](https://github.com/anasahmed07/Outlook-Classic-MCP) at 0.3.2.

### Merged upstream pull requests

- #19 Fix timezone shift when writing dates to Outlook (`to_com`).
- #11 Locale-independent date filtering (DASL), overlap filtering for multi-day
  events, `truncated` flag on `list_events`, `include_attachments` on
  `get_mail`, tolerant attachment reads, coordinated Outlook relaunch and a
  45 s call timeout.
- #16 Resolve recipients before saving a draft.
- #18 `account` and `signature` parameters on `outlook_send_mail`.

### Mail

- `outlook_reply_mail` and `outlook_forward_mail` take `save_only` and save a
  threaded draft (quoted history kept) instead of sending.
- New `outlook_send_draft` sends an existing draft after review.
- Replies and forwards take `to` / `cc` / `bcc` (added to the reply's own),
  `subject`, `importance`, `signature`; forwards take `attachments`, and `to`
  is optional for drafts.
- Reply/forward text is inserted inside `<body>` / `WordSection1` as
  `MsoNormal` paragraphs, so it uses the message's styles. Previously it was
  placed before `<html>`, and plain-text replies to HTML mail replaced the
  whole HTML body (flattening the quoted thread).
- Signatures: `signature=true`/`"default"` uses the account's default from the
  Outlook profile (new-mail default for new mail, reply/forward default for
  replies), or pass a signature name. Images in signatures are embedded as
  hidden inline attachments. A default signature Outlook already inserted
  into a reply is not added twice. New `outlook_list_signatures`.
- Recipients are resolved before saving or sending. Sending with unresolved
  recipients is refused with the names listed; drafts are saved and the names
  reported in `notes`.
- Attachment paths are validated before the item is created.

### Calendar

- `all_day` on create/update; `end` is optional for all-day events and the
  last day is inclusive.
- `send_invites=false` saves a meeting without sending invitations.
- `optional_attendees`, `busy_status`; `reminder_minutes=null` turns the
  reminder off.
- `update_event`: `send_update` (organizer only) sends the update to
  attendees; `add_attendees`, `add_optional_attendees`, `remove_attendees`,
  `reminder_minutes`, `busy_status`, `all_day`.
- `delete_event` sends a cancellation for meetings you organize
  (`send_cancellation=false` to skip). The old docs claimed this happened
  already; it didn't.
- `occurrence_start` on get/update/delete/respond targets one occurrence of a
  recurring series. Updating or deleting a whole series requires
  `apply_to_series=true`.
- `respond_event` takes a `comment`.
- `list_events` items carry `recurrence` (`none` / `series` / `occurrence` /
  `exception`) and stop early once past the range.

### Safety and configuration

- `OUTLOOK_MCP_DRAFT_ONLY=1`: nothing is sent. Compose tools save drafts,
  `send_draft` is refused, meeting invites/updates are saved unsent,
  cancellations are refused, meeting responses are left in Drafts.
- `OUTLOOK_MCP_ALLOWED_PATHS`: extra folders allowed for attachments and
  saved attachments.
- `OUTLOOK_MCP_SIGNATURES_DIR`: override the signatures folder.

### Fixes

- DASL date filters now convert local time to UTC (Outlook evaluates DASL
  dates in UTC), so `since` / `until` and calendar ranges are no longer off by
  the UTC offset.

### Project

- Install from this repository (plugin, clone + `install.bat`, or
  `uvx --from git+...`); the upstream PyPI publish workflow was removed.
- Test suite runs on Linux too (COM is mocked; `pythoncom` stubbed); CI on
  Windows and Linux.
