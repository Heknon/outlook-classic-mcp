# Changelog

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
