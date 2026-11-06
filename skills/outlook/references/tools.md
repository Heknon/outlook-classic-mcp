# Tool reference

Every `outlook_*` tool, with parameters, defaults, return shape, and notes on chaining. Skim the table of contents, then jump to the tools you need.

## Contents

- [Mail](#mail) — list_mails, search_mails, search_mails_by_person, get_mail, get_thread, send_mail, reply_mail, forward_mail, send_draft, list_signatures, move_mail, delete_mail, mark_mail, save_attachments
- [Folders](#folders) — list_folders, list_mailboxes, create_folder
- [Calendar](#calendar) — list_events, get_event, create_event, update_event, delete_event, respond_event, get_free_busy, find_meeting_times
- [Contacts](#contacts) — list_contacts, search_contacts, get_contact, resolve_name
- [Tasks](#tasks) — list_tasks, create_task, complete_task
- [Categories](#categories) — list_categories, set_category
- [Rules](#rules) — list_rules, toggle_rule
- [Out-of-Office](#out-of-office) — get_out_of_office
- [Account](#account) — whoami
- [Common return-field glossary](#common-return-field-glossary)

---

## Mail

### `outlook_list_mails`

List mail items from a folder, newest first. Read-only.

| Param            | Type      | Default     | Notes |
| ---------------- | --------- | ----------- | ----- |
| `folder`         | string    | `"inbox"`   | Well-known name or path. See SKILL.md → Folder references. |
| `limit`          | int 1–100 | `25`        | Max items to return. |
| `offset`         | int ≥0    | `0`         | Skip this many before returning. Use with the returned `next_offset` to paginate. |
| `unread_only`    | bool      | `false`     | If true, only unread mails. |
| `since`          | ISO-8601  | `null`      | Lower bound on `ReceivedTime`. |
| `until`          | ISO-8601  | `null`      | Upper bound on `ReceivedTime`. |
| `from_address`   | string    | `null`      | **Substring** match on sender. See gotcha re: `EX:/O=...` addresses. |
| `response_format`| `markdown`/`json` | `markdown` | Use `json` to extract `entry_id`s. |

**Returns** (`json` shape): `{ folder, count, offset, limit, items: [...], has_more, next_offset }`. Each item has: `entry_id, subject, from, from_address, to, received, unread, has_attachments, importance, preview` (200-char body excerpt).

### `outlook_search_mails`

Search a single folder by subject/body, subject-only, sender, or raw DASL. Read-only.

| Param            | Type     | Default          | Notes |
| ---------------- | -------- | ---------------- | ----- |
| `query`          | string   | required         | Search words (ALL must match, any order) — or a DASL @SQL filter when `scope='dasl'`. |
| `folder`         | string   | `"inbox"`        | Where to search. |
| `scope`          | enum     | `"subject_body"` | `subject_body` (default), `subject`, `from`, or `dasl`. |
| `limit`          | int 1–100| `25`             | |
| `response_format`| str      | `markdown`       | |

**Returns**: `{ query, scope, folder, count, items: [...] }`. Items have the same summary shape as `list_mails`.

Multi-word queries match items containing **all** the words (not the exact phrase), so `"teams not working"` finds "MESP-1 teams is not working". `scope='from'` matches display name, raw address, **and** the real SMTP address (works for Exchange senders too).

`scope='dasl'` is for power use — pass a complete `@SQL=...` filter and the server applies it raw. Only reach for this when subject_body/subject/from can't express what the user wants.

### `outlook_search_mails_by_person`

Find mails exchanged with someone given just their name ("emails from Dana", "what did I send Yossi last week").

| Param       | Type      | Default | Notes |
| ----------- | --------- | ------- | ----- |
| `person`    | string    | required | Name, partial name, alias or address — resolved like Ctrl+K first. |
| `direction` | enum      | `"any"` | `from` (they sent it), `to` (they were on To/CC), `any`. |
| `folders`   | list[str] | inbox + sent | Inbox only for `from`. |
| `since`     | ISO-8601  | `null`  | |
| `limit`     | int 1–100 | `25`    | |
| `response_format` | str | `markdown` | |

**Returns**: `{ person, direction, folders, count, items: [...] }`, items as in `list_mails` plus `folder`. If the name matches several people, nothing is searched: `{ status: "ambiguous", candidates, count: 0, items: [] }` — ask which person, then call again with their `smtp_address`.

### Searching several mailboxes (`search_in`, `mailboxes`)

Both `outlook_search_mails` and `outlook_search_mails_by_person` accept:

| Param       | Type      | Default   | Notes |
| ----------- | --------- | --------- | ----- |
| `search_in` | enum      | `"folder"`| `folder` (only `folder`/`folders`), `mailbox` (every folder of your mailbox), `archive` (online/in-place archive), `all` (every mailbox in the profile incl. shared and PSTs; not public folders). |
| `mailboxes` | list[str] | `null`    | Only these mailboxes (case-insensitive part of the name from `outlook_list_mailboxes`); every folder in each. |
| `include_deleted` | bool | `false`  | Include Deleted Items / Junk. |

Multi-mailbox results carry `folder` ("Mailbox/Inbox/Projects"), `search_in`, `folders_searched` and, if some folders couldn't be searched, `folder_errors`.

### `outlook_get_mail`

Fetch the body, all headers, and the attachment manifest for one mail. Read-only.

| Param            | Type   | Default | Notes |
| ---------------- | ------ | ------- | ----- |
| `entry_id`       | string | required | From a list/search result. |
| `include_body`   | bool   | `true`   | If false, omits `body`. Useful when you only need metadata. |
| `include_html`   | bool   | `false`  | Adds the raw `html_body`. Usually huge — leave off unless you specifically need the markup. |
| `max_body_chars` | int ≥0 | `10000`  | Body truncation cap; `0` = unlimited. |
| `response_format` | str | `markdown` | |

**Returns**: `{ entry_id, conversation_id, subject, from, from_address, to, cc, bcc, received, sent, unread, importance, categories, attachments: [{index, filename, size_bytes}], body }` plus `body_truncated`/`body_total_chars` when the cap was hit (re-call with a higher `max_body_chars` to read more) and `html_body` when `include_html=true`.

`attachments[].index` is **1-indexed**; pass it to `save_attachments` to save a single file.

### `outlook_get_thread`

Every message in the conversation that `entry_id` belongs to, oldest first, in one call. Uses Outlook's conversation view (all folders, incl. Sent Items and drafts); falls back to matching the conversation topic in Inbox + Sent Items (`source: "topic_search"`).

| Param            | Type   | Default | Notes |
| ---------------- | ------ | ------- | ----- |
| `entry_id`       | string | required | Any message in the thread. |
| `strip_quoted`   | bool   | `true`  | Cut each message's quoted history. |
| `max_body_chars` | int    | `4000`  | Per message; `0` = no limit. |
| `limit`          | int    | `50`    | Keeps the most recent messages. |

**Returns**: `{ topic, count, total_in_thread, source, messages: [{entry_id, when, from, to, cc, subject, folder, is_draft, unread, attachment_count, body, quoted_history_removed?, body_truncated?}] }`.

### `outlook_send_mail`

Compose a new mail and send it, or save it to Drafts. Has external side effect unless `save_only=true`.

| Param          | Type      | Default   | Notes |
| -------------- | --------- | --------- | ----- |
| `to`           | list[str] | required  | One or more SMTP addresses or resolvable names. |
| `subject`      | string    | required  | |
| `body`         | string    | required  | Plain text unless `html=true`. |
| `cc`           | list[str] | `null`    | |
| `bcc`          | list[str] | `null`    | |
| `html`         | bool      | `false`   | When true, `body` is HTML. |
| `attachments`  | list[str] | `null`    | Absolute paths under the user profile (or `OUTLOOK_MCP_ALLOWED_PATHS`). |
| `importance`   | enum      | `"normal"`| `low` / `normal` / `high`. |
| `save_only`    | bool      | `false`   | **Save to Drafts instead of sending.** |
| `account`      | string    | `null`    | Send from (and save into the Drafts of) this account — SMTP address or display name. |
| `signature`    | bool/str  | `null`    | `true`/`"default"` = the account's default signature for new mail; or a name from `outlook_list_signatures`. |

**Returns** (sent): `{ status: "sent", subject, to, cc, bcc }`. (Drafts): `{ status: "saved_to_drafts", entry_id, subject, to, cc, bcc }`. Either may carry `signature` (the name used) and `notes` (e.g. draft-only mode, unresolved recipients).

Always confirm the recipient list and subject with the user before sending unless they have explicitly authorized you to send. Recipients that can't be resolved block a send with an error; a draft is still saved and the names are listed in `notes`.

### `outlook_reply_mail`

Reply (or reply-all) to an existing mail. Uses Outlook's own Reply, so the message stays in the conversation thread and the quoted original (with its formatting) is kept below your text. Has external side effect unless `save_only=true`.

| Param         | Type      | Default | Notes |
| ------------- | --------- | ------- | ----- |
| `entry_id`    | string    | required | The mail being replied to. |
| `body`        | string    | required | Your reply text, placed above the quoted original. |
| `reply_all`   | bool      | `false`  | If true, includes the original To/CC list. |
| `html`        | bool      | `false`  | Treat `body` as an HTML fragment. |
| `attachments` | list[str] | `null`   | |
| `to` / `cc` / `bcc` | list[str] | `null` | Extra recipients, added to the reply's own. |
| `subject`     | string    | `null`   | Replace the subject (default keeps `RE: ...`). |
| `importance`  | enum      | `null`   | `low` / `normal` / `high`. |
| `signature`   | bool/str  | `null`   | `true` = default reply/forward signature (skipped if Outlook already added one), or a signature name. |
| `save_only`   | bool      | `false`  | **Save the threaded reply to Drafts instead of sending.** |

**Returns**: `{ status: "sent" | "saved_to_drafts", in_reply_to, reply_all, subject, entry_id? , notes? }`.

To stage a reply for review: `outlook_reply_mail(..., save_only=true)`, tell the user, and once they approve call `outlook_send_draft(entry_id)` (or let them send it from Outlook).

### `outlook_forward_mail`

Forward an existing mail (with its attachments) and an optional note above it. Has external side effect unless `save_only=true`.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `entry_id` | string    | required | |
| `to`       | list[str] | `null`   | Required unless `save_only=true`. |
| `body`     | string    | `""`     | Optional note above the forwarded content. |
| `cc` / `bcc` | list[str] | `null` | |
| `html`     | bool      | `false`  | |
| `attachments` | list[str] | `null` | Extra files to attach. |
| `subject`  | string    | `null`   | Replace the subject (default keeps `FW: ...`). |
| `importance` | enum    | `null`   | |
| `signature` | bool/str | `null`   | As for replies. |
| `save_only` | bool     | `false`  | Save to Drafts instead of sending. |

**Returns**: `{ status, forwarded, to, subject, entry_id?, notes? }`.

### `outlook_send_draft`

Send a draft that is already in Drafts (from a `save_only` result or found with `outlook_list_mails(folder="drafts")`). Confirm with the user first. Refused in draft-only mode, for items already sent, and for drafts with no or unresolvable recipients.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

**Returns**: `{ status: "sent", subject, to, cc, bcc }`.

### `outlook_list_signatures`

List the user's saved Outlook signatures and which are the defaults. No parameters.

**Returns**: `{ folder, signatures: [names], default_new, default_reply_forward }`.

### `outlook_move_mail`

Move a mail to another folder.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `target_folder` | string | required | Well-known name or path. |

**Returns**: `{ status: "moved", new_entry_id, folder }`.

The `entry_id` changes when an item moves stores. **Use the returned `new_entry_id`** if you need to act on the moved item again.

### `outlook_delete_mail`

Soft-delete (moves to Deleted Items). Reversible by the user from Outlook.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

**Returns**: `{ status: "deleted", subject, entry_id }`.

### `outlook_mark_mail`

Toggle read state and/or follow-up flag.

| Param      | Type | Default | Notes |
| ---------- | ---- | ------- | ----- |
| `entry_id` | string | required | |
| `read`     | bool/null | `null` | `true` = mark read, `false` = mark unread, `null` = no change. |
| `flagged`  | bool/null | `null` | `true` = flag for follow-up, `false` = clear flag, `null` = no change. |

**Returns**: `{ status: "updated", entry_id, unread, flagged }`.

### `outlook_save_attachments`

Save one or all attachments from a mail to a local directory.

| Param              | Type    | Default | Notes |
| ------------------ | ------- | ------- | ----- |
| `entry_id`         | string  | required | |
| `output_dir`       | string  | required | Absolute path under user profile. Created if missing. |
| `attachment_index` | int ≥1  | `null`  | 1-indexed. Omit to save all. |

**Returns**: `{ status: "saved", count, files: [absolute paths], output_dir }`.

---

## Folders

### `outlook_list_folders`

Walk the folder tree under a root.

| Param            | Type    | Default | Notes |
| ---------------- | ------- | ------- | ----- |
| `root`           | string  | `null`  | Folder to start from. Default = the default mailbox root. |
| `max_depth`      | int 1–10| `4`     | How deep to walk. |
| `response_format`| str     | `markdown` | |

**Returns**: `{ count, items: [{name, path, item_count, unread_count, default_item_type}, ...] }`. The `path` strings are exactly what you pass back as a `folder` parameter elsewhere.

### `outlook_list_mailboxes`

Mailboxes open in Outlook. No parameters.

**Returns**: `{ count, mailboxes: [{name, kind, is_default}] }`, `kind` one of `primary`, `archive`, `shared`, `delegate`, `pst_or_other`, `public_folders`. An archive is recognized by its name ("Online Archive - ...", "In-Place Archive - ...", and the German/Spanish/Hebrew equivalents); if one isn't detected, use its name with `mailboxes=[...]`.

### `outlook_create_folder`

Create a sub-folder under a parent.

| Param    | Type   | Default   | Notes |
| -------- | ------ | --------- | ----- |
| `name`   | string | required  | New folder name. |
| `parent` | string | `"inbox"` | Where to put it. |

**Returns**: `{ name, path, entry_id }`.

---

## Calendar

Times are local wall-clock ISO-8601 strings (an explicit offset is converted to local). For recurring meetings, every occurrence returned by `outlook_list_events` carries the **series** `entry_id`; its `recurrence` field is `occurrence` (or `exception` for a moved one). To act on one occurrence pass `occurrence_start` = that item's `start`, exactly as listed. Changing or deleting a whole series (`recurrence: "series"`, no `occurrence_start`) is refused unless `apply_to_series=true`.

### `outlook_list_events`

List calendar events overlapping a date range, including expanded recurring instances.

| Param                 | Type    | Default        | Notes |
| --------------------- | ------- | -------------- | ----- |
| `start`               | ISO-8601| now            | |
| `end`                 | ISO-8601| `start + 14d`  | |
| `limit`               | int 1–200| `50`          | |
| `include_recurrences` | bool    | `true`         | If false, only the master entries — usually you want true. |
| `mailbox`             | string  | `null`         | Someone else's calendar (address or name). Needs them to have shared it with you; otherwise use free/busy. |
| `response_format`     | str     | `markdown`     | |

**Returns**: `{ start, end, count, items: [...], truncated, mailbox? }`. Items: `entry_id, subject, start, end, location, organizer, is_recurring, recurrence ("none" | "series" | "occurrence" | "exception"), all_day, preview`.

### `outlook_get_event`

Full event detail, including attendees and their RSVP status.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |
| `occurrence_start` | ISO-8601 | `null` | One occurrence of a series. |
| `mailbox`  | string | `null` | The calendar owner, for events listed with `mailbox`. |
| `response_format` | str | `markdown` | |

**Returns**: summary fields + `body, attendees: [{name, address, type, response}], reminder_minutes, categories`.

### `outlook_create_event`

Create an appointment, all-day event, or meeting. **With attendees, invitations are sent on success unless `send_invites=false`.**

| Param               | Type           | Default | Notes |
| ------------------- | -------------- | ------- | ----- |
| `subject`           | string         | required | |
| `start`             | ISO-8601       | required | A date (`2026-12-24`) is fine for all-day events. |
| `end`               | ISO-8601       | `null`  | Required unless `all_day`; for all-day it is the last day, inclusive (default: one day). |
| `location`          | string         | `null`  | |
| `body`              | string         | `null`  | |
| `attendees`         | list[str]      | `null`  | Required attendees — makes this a meeting. |
| `optional_attendees`| list[str]      | `null`  | |
| `all_day`           | bool           | `false` | |
| `send_invites`      | bool           | `true`  | `false` saves the meeting unsent; the user opens it in Outlook and clicks Send. |
| `busy_status`       | enum           | `null`  | `free` / `tentative` / `busy` / `out_of_office` / `working_elsewhere`. |
| `reminder_minutes`  | int 0–10080    | `15`    | `null` = no reminder. |
| `recurrence`        | Recurrence obj | `null`  | See SKILL.md → Recurrence. |
| `is_online_meeting` | bool           | `false` | Reserved; ignored by classic Outlook COM. |

**Returns**: `{ status: "created", entry_id, subject, start, end, all_day, invites_sent, notes? }`.

Confirm attendee list, times, and recurrence with the user before calling.

### `outlook_update_event`

Update an event. Only non-null fields are written. Does **not** modify recurrence — for cadence changes, delete and recreate.

| Param      | Type     | Default | Notes |
| ---------- | -------- | ------- | ----- |
| `entry_id` | string   | required | |
| `subject`, `location`, `body` | string | `null` | |
| `start` / `end` | ISO-8601 | `null` | |
| `all_day`  | bool     | `null`  | Turn all-day on/off (with `start`, snaps to whole days). |
| `reminder_minutes` | int | `null` | |
| `busy_status` | enum  | `null`  | |
| `add_attendees` / `add_optional_attendees` | list[str] | `null` | |
| `remove_attendees` | list[str] | `null` | Address or display name. |
| `send_update` | bool  | `false` | For meetings you organize: send the update to attendees. Without it only your calendar changes. |
| `occurrence_start` | ISO-8601 | `null` | Change just this occurrence. |
| `apply_to_series` | bool | `false` | Required to change a whole recurring series. |

**Returns**: `{ status: "updated", entry_id, update_sent, removed_attendees?, notes? }`.

### `outlook_delete_event`

Delete an event. For a meeting you organize, a cancellation is sent to attendees first (`send_cancellation=false` deletes it from your calendar only). As an attendee, the meeting is just removed — use `outlook_respond_event` to decline if the organizer should know.

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |
| `send_cancellation` | bool | `true` | |
| `occurrence_start` | ISO-8601 | `null` | Delete (cancel) just this occurrence. |
| `apply_to_series` | bool | `false` | Required to delete a whole series. |

**Returns**: `{ status: "deleted", subject, entry_id, cancellation_sent, notes? }`.

### `outlook_respond_event`

Respond to a meeting invite.

| Param           | Type   | Default | Notes |
| --------------- | ------ | ------- | ----- |
| `entry_id`      | string | required | |
| `response`      | enum   | required | `accept` / `tentative` / `decline`. |
| `send_response` | bool   | `true`   | Set false to record locally without emailing the organizer. |
| `comment`       | string | `null`   | Note included in the response. |
| `occurrence_start` | ISO-8601 | `null` | Respond to one occurrence. |

**Returns**: `{ status: "responded" | "response_saved_to_drafts", response }`.

### `outlook_get_free_busy`

When people are busy, from Exchange free/busy. Works for anyone in the organization without calendar permission; shows no meeting details.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `people`   | list[str] | required | Addresses or names (1–30). |
| `start` / `end` | ISO-8601 | now / +1 day | At most 62 days. |
| `interval_minutes` | int 5–60 | `30` | Resolution. |
| `include_me` | bool | `false` | |

**Returns**: `{ start, end, interval_minutes, people: [{name, busy: [{start, end, status}]}], problems? }`. `status` is `busy`, `tentative`, `out_of_office`, `working_elsewhere` or `unknown`. People whose free/busy can't be read (external, unresolvable) are under `problems`.

### `outlook_find_meeting_times`

Scheduling assistant: slots when you and all required attendees are free.

| Param      | Type      | Default | Notes |
| ---------- | --------- | ------- | ----- |
| `attendees` | list[str] | required | Required attendees. |
| `duration_minutes` | int | `30` | |
| `optional_attendees` | list[str] | `null` | Listed per slot under `optional_busy` if they can't make it. |
| `start` / `end` | ISO-8601 | now / +7 days | |
| `day_start` / `day_end` | HH:MM | `09:00` / `17:00` | Local working hours; the meeting must fit inside. |
| `working_days` | list[str] | Mon–Fri | e.g. `["sun","mon","tue","wed","thu"]`. |
| `step_minutes` | int | `30` | Start-time grid: 5, 10, 15, 20, 30 or 60. |
| `include_me` | bool | `true` | Require the user to be free too. |
| `tentative_is_free` | bool | `false` | |
| `max_results` / `max_per_day` | int | `10` / `3` | |

**Returns**: `{ duration_minutes, searched, required, optional, count, slots: [{start, end, optional_busy?, tentative?}], problems?, notes? }`. "Working elsewhere" counts as available; unknown free/busy (beyond published data) counts as busy.

Then create the meeting with `outlook_create_event(attendees=..., start=slot.start, end=slot.end)`.

---

## Contacts

### `outlook_list_contacts`

List saved contacts from **every contact folder in every store**, sorted by full name within each folder.

| Param            | Type     | Default | Notes |
| ---------------- | -------- | ------- | ----- |
| `limit`          | int 1–200| `50`    | |
| `offset`         | int ≥0   | `0`     | |
| `response_format`| str      | `markdown` | |

**Returns**: `{ count, offset, items: [...], has_more }`. Items: `entry_id, full_name, email, company, job_title, mobile, folder`.

On corporate accounts the personal contact folders are often nearly empty — colleagues live in the **directory** (Global Address List), which this tool does not list. Use `search_contacts` to find people.

### `outlook_search_contacts`

Word search across saved contacts (name, email, company, job title) **and the org directory (GAL)**. All query words must match.

| Param               | Type    | Default | Notes |
| ------------------- | ------- | ------- | ----- |
| `query`             | string  | required | e.g. `"anas shaikh"` matches "Anas Ahmed Shaikh". |
| `limit`             | int 1–100| `25`   | |
| `include_directory` | bool    | `true`  | Also scan the Exchange Global Address List. A few seconds on large directories. |
| `response_format`   | str     | `markdown` | |

**Returns**: `{ query, count, items, searched_directory }`. Each item has `source: "contacts" | "directory"`. Directory items carry `full_name, email (SMTP), company, job_title, mobile` but **no `entry_id`** — they aren't Outlook items, so don't pass them to `get_contact`.

### `outlook_get_contact`

Full contact record (saved contacts only — needs an `entry_id`). Returns the summary fields plus `business_phone, home_phone, address, notes`.

### `outlook_resolve_name`

Check names, like typing a name in To: and pressing **Ctrl+K**: get a person's email address from a full or partial name, alias or address. Searches your contacts and the organization's address book.

| Param  | Type   | Default | Notes |
| ------ | ------ | ------- | ----- |
| `name` | string | required | Full name, part of a name, alias (`dcohen`) or address. |
| `max_candidates` | int 1–50 | `10` | |

**Returns** one of:
- `{ status: "resolved", resolved: true, person: {display_name, smtp_address, job_title, department, office, phone, alias}, display_name, smtp_address }`
- `{ status: "ambiguous", resolved: false, candidates: [{display_name, smtp_address, job_title, department, office, source}], note }` — **ask the user which one**; don't guess.
- `{ status: "not_found", resolved: false, note }`

---

## Tasks

### `outlook_list_tasks`

List tasks from the default Tasks folder, sorted by due date.

| Param                | Type     | Default | Notes |
| -------------------- | -------- | ------- | ----- |
| `limit`              | int 1–200| `50`    | |
| `include_completed`  | bool     | `false` | Default hides done tasks. |
| `response_format`    | str      | `markdown` | |

**Items**: `entry_id, subject, due_date, start_date, complete, percent_complete, importance, status`.

### `outlook_create_task`

| Param        | Type     | Default   | Notes |
| ------------ | -------- | --------- | ----- |
| `subject`    | string   | required  | |
| `due_date`   | ISO-8601 | `null`    | |
| `body`       | string   | `null`    | |
| `importance` | enum     | `"normal"`| low/normal/high. |
| `reminder`   | ISO-8601 | `null`    | Sets a reminder time. |

**Returns**: `{ status: "created", entry_id, subject }`.

### `outlook_complete_task`

| Param      | Type   | Default | Notes |
| ---------- | ------ | ------- | ----- |
| `entry_id` | string | required | |

Marks the task 100% complete. Returns `{ status: "completed", entry_id }`.

---

## Categories

### `outlook_list_categories`

Returns the color categories defined in the user's Outlook profile: `{ count, items: [{name, color}, ...] }`. Categories are profile-wide, not per-folder.

### `outlook_set_category`

Replace the categories on any item (mail, event, task).

| Param        | Type   | Default | Notes |
| ------------ | ------ | ------- | ----- |
| `entry_id`   | string | required | |
| `categories` | string | required | **Comma-separated names.** Empty string clears all. e.g. `"Important"` or `"Work, Follow-up"`. |

This **replaces** existing categories rather than adding to them. To add `Foo` to an item that already has `Bar`, send `"Bar, Foo"`. Get the current value first via `get_mail` / `get_event` if needed.

---

## Rules

### `outlook_list_rules`

Returns the user's mail rules with their on/off state: `{ count, items: [{index, name, enabled}] }`.

### `outlook_toggle_rule`

| Param       | Type   | Default | Notes |
| ----------- | ------ | ------- | ----- |
| `rule_name` | string | required | **Exact** name from `list_rules`. |
| `enabled`   | bool   | required | `true` to enable, `false` to disable. |

This change is live the moment it's saved — no staging buffer. Confirm the rule name with the user before calling.

---

## Out-of-Office

### `outlook_get_out_of_office`

Reports whether OOO auto-reply is currently on. Returns `{ out_of_office: bool, status: "on"|"off" }`, or `{ out_of_office: null, status: "unknown", note: ... }` if the property isn't readable on this profile.

There is **no tool to enable, disable, or schedule OOO**. Tell users to manage it via Outlook → File → Automatic Replies.

---

## Account

### `outlook_whoami`

Returns the bound user, the account list, and the user's timezone: `{ current_user, accounts: [{display_name, smtp_address, user_name, account_type}, ...], local_time, timezone, utc_offset }`. Useful as a sanity check when the user has multiple mailboxes, and as the authority on what timezone all returned datetimes are in.

---

## Common return-field glossary

- `entry_id` — opaque, stable handle for an item. Pass back verbatim. Becomes invalid on delete; changes on cross-store move.
- `conversation_id` — groups mails in a thread. Same value across replies/forwards in one conversation.
- `from` — display name of the sender.
- `from_address` — sender address. **For Exchange senders, this is an `EX:/O=...` distinguished name, not SMTP.** Match by substring.
- `received` / `sent` / `start` / `end` / `due_date` — ISO-8601 strings **in the user's local timezone with explicit offset** (e.g. `2026-06-10T16:33:22+05:00`). Present as-is; never convert to another timezone.
- `unread` — bool. Note `mark_mail` returns `unread` (not `read`).
- `importance` — integer (0=low, 1=normal, 2=high).
- `categories` — comma-separated string of category names; empty string = none.
- `preview` — 200-char body excerpt. Not a substitute for `get_mail` / `get_event` when you need the full body.
- `has_more` / `next_offset` — pagination signals on list endpoints.
