---
name: outlook
description: >-
  Use Microsoft Outlook end-to-end — read, search, triage, and send mail; manage calendar events and meeting invites; work with contacts, tasks, color categories, and mail rules; check Out-of-Office. Trigger this skill whenever the user asks to do anything with their inbox, email, calendar, meetings, contacts, or to-do list AND tools whose names start with `outlook_` are available (e.g. `outlook_list_mails`, `outlook_send_mail`, `outlook_create_event`). Trigger even when the user doesn't say "Outlook" — phrases like "check my inbox", "what's on my calendar this week", "draft a reply to that email about Y", "schedule a meeting with Sarah Friday at 3", "remind me to follow up next Monday", "move that thread to a folder", "find Bob's phone number", or "is my OOO still on?" all qualify. The skill teaches the right way to chain these tools: how to refer to folders, how item handles work, how to format dates, how to recur an event, and which calls have side effects you must confirm first.
---

# Outlook

You have tools whose names begin with `outlook_` that operate on the user's actual Microsoft Outlook account. They read and write live data — sent mail goes out, deleted items move to Deleted Items, accepted invites notify organizers. Treat write operations the way you would treat any user-visible action: confirm intent first.

## When to use this skill

Activate whenever the user wants to do something with their Outlook data. The user does not need to say "Outlook." Triggers include:

- **Mail**: "any new emails from X?", "draft a reply to that thread about Y", "forward Bob's note to the team", "clean out my inbox", "save the attachment from Sarah's email"
- **Calendar / meetings**: "what do I have Friday?", "schedule 30 min with Alice next Tuesday", "make that meeting recurring weekly", "decline the 4pm", "what time is the all-hands?"
- **Contacts**: "what's Carlos's email?", "phone number for the legal team contact"
- **Tasks**: "remind me to send the report Monday", "mark the budget task done"
- **Categories / rules / OOO**: "tag that email as Important", "is my OOO still on?", "turn off the rule that auto-files Jira mail"

If `outlook_*` tools are not available in the session, the underlying integration isn't installed yet. Don't guess or improvise — instead, read `references/setup.md` and walk the user through installing it (Windows + classic Outlook + a one-line install + restart their MCP client). After install they restart the client and the tools appear.

## Tool catalog

All tools are prefixed `outlook_`. Memorize the categories; consult `references/tools.md` for exact parameters and return shapes the first time you reach for one.

| Category       | Tools |
| -------------- | ----- |
| Mail           | `list_mails`, `search_mails`, `get_mail`, `send_mail`, `reply_mail`, `forward_mail`, `move_mail`, `delete_mail`, `mark_mail`, `save_attachments` |
| Folders        | `list_folders`, `create_folder` |
| Calendar       | `list_events`, `get_event`, `create_event`, `update_event`, `delete_event`, `respond_event` |
| Contacts       | `list_contacts`, `search_contacts` (incl. org directory), `get_contact`, `resolve_name` |
| Tasks          | `list_tasks`, `create_task`, `complete_task` |
| Categories     | `list_categories`, `set_category` |
| Rules          | `list_rules`, `toggle_rule` |
| Out-of-Office  | `get_out_of_office` (read-only) |
| Account        | `whoami` |

## Conventions you must internalize

These rules are the same across every tool. Internalize them and you'll rarely need to peek at `references/tools.md`.

### Folder references

Any parameter named `folder`, `target_folder`, `parent`, or `root` accepts:

- A **well-known name** (case-insensitive): `inbox`, `sent`, `drafts`, `deleted` (alias `trash`), `outbox`, `junk` (alias `spam`), `calendar`, `contacts`, `tasks`, `notes`.
- A **slash path** under the default mailbox: `Inbox/Projects/Acme`.
- A **store-qualified path** when the user has more than one mailbox or PST: `Mailbox - you@example.com/Inbox/Projects/Acme`.

When unsure of the exact path, call `outlook_list_folders` first and use the path string it returns verbatim.

### Item identity: `EntryID`

Every mail, event, contact, and task has a stable `entry_id` string. List/search tools return it on every item; pass that exact string back to detail/edit/delete/reply/move tools. **Never invent an `entry_id`. Never paste a subject line into an `entry_id` field.**

If the user refers to an item by description ("the email from Bob about Q3"), first find it (e.g. `search_mails`), then act on the returned `entry_id`.

`entry_id`s become invalid if the item is deleted; moving an item between stores produces a new ID. If you get an "Item not found" error, re-list to refresh.

### Dates and times

All date/time parameters are **ISO-8601** strings: `2026-04-25T14:30:00`. Without a timezone suffix the value is interpreted as the user's local time, which is what the user expects.

Resolve relative phrasing ("Friday at 3", "next Monday morning", "tomorrow") yourself before calling — the tools don't parse natural language. Use today's date from the conversation context as the anchor.

### `response_format`

Most read tools accept `response_format='markdown'` (default; pretty for the user) or `response_format='json'` (structured for you to parse). Pick **`json`** when you need to extract a field (almost always an `entry_id`) to chain into another call. Pick **`markdown`** when the result is the final answer to the user.

### Recurrence (calendar)

`outlook_create_event` accepts an optional `recurrence` object:

```json
{
  "type": "weekly",         // "daily" | "weekly" | "monthly" | "yearly"
  "interval": 1,            // every N units; default 1
  "occurrences": 10,        // OR
  "end_date": "2026-12-31"  // ISO date — provide one or the other; omit both for indefinite
}
```

`update_event` does not change recurrence. If the user wants to alter a recurring series' cadence, the simplest path is delete + recreate.

### Filesystem paths

`attachments=` (on `send_mail`/`reply_mail`/`forward_mail`) and `output_dir=` (on `save_attachments`) require **absolute paths under the user's profile directory** (`C:\Users\<them>\...`). This sandbox is enforced; you can't bypass it from a tool call. If the user really needs a path elsewhere (a network share, D:\, etc.), they have to set `OUTLOOK_MCP_ALLOW_ANY_PATH=1` in their environment — tell them, don't try to work around it.

### Read tools are free; write tools have side effects

Read freely:
`list_mails`, `search_mails`, `search_mails_by_person`, `get_mail`, `get_thread`, `list_folders`, `list_events`, `get_event`, `list_contacts`, `search_contacts`, `get_contact`, `resolve_name`, `list_tasks`, `list_categories`, `list_rules`, `get_out_of_office`, `list_signatures`, `list_mailboxes`, `get_free_busy`, `find_meeting_times`, `whoami`.

Confirm before calling (these change shared state or send messages):
`send_mail`, `reply_mail`, `forward_mail`, `send_draft`, `delete_mail`, `move_mail`, `mark_mail`, `save_attachments`, `create_event` (especially with attendees — that sends a meeting invite immediately), `update_event`, `delete_event`, `respond_event` (with `send_response=true`), `create_folder`, `create_task`, `complete_task`, `set_category`, `toggle_rule`.

Staging (draft) options — prefer these whenever the user says "draft" or hasn't explicitly approved sending:
- `outlook_send_mail`, `outlook_reply_mail` and `outlook_forward_mail` all take `save_only=true`: the message is saved to Drafts and its `entry_id` returned. Replies and forwards stay properly threaded with the quoted history intact. Tell the user it's in Drafts.
- `outlook_send_draft(entry_id)` sends a saved draft once the user approves it — no need to recreate it.
- `outlook_create_event(..., send_invites=false)` saves a meeting in the calendar without sending the invitations.
- `outlook_update_event` changes only the user's calendar unless `send_update=true`.
- `outlook_respond_event(..., send_response=false)` records your local accept/decline without emailing the organizer.

**By default the server is in draft-only mode**: nothing can be sent. Compose tools always save drafts (the result has `status: "saved_to_drafts"` and a note), `send_draft` is refused, and invites/updates/cancellations are not sent. Don't try to work around it — tell the user the draft is waiting in Outlook. Only if the user has set `OUTLOOK_MCP_ALLOW_SEND=1` can things be sent.

### Threads

To summarize a conversation or reply to a long one, call `outlook_get_thread(entry_id)` once instead of `outlook_get_mail` on each message. It includes the user's own replies from Sent Items, oldest first, with quoted history removed. Use `strip_quoted=false` only if something seems missing.

### Writing in Hebrew

Write Hebrew drafts as normal plain text in `body`; the server makes the paragraphs right-to-left. Don't add your own HTML or direction marks.

### Searching beyond the Inbox

`outlook_search_mails` / `outlook_search_mails_by_person` only search one folder (Inbox by default). If the user says "anywhere", "in my archive", "in the Support mailbox", or a search of the Inbox finds nothing, widen it: `search_in="mailbox"` (all of their folders), `"archive"` (online archive), `"all"` (every mailbox incl. shared and PSTs), or `mailboxes=[...]` with names from `outlook_list_mailboxes`. Each result has its `folder` path. Searching everything is slower — start narrow.

### People: names to addresses (Ctrl+K)

- When the user names someone ("email Dana", "invite Yossi"), get the address with `outlook_resolve_name` — it works like Ctrl+K and accepts partial names and aliases. If the status is `ambiguous`, show the candidates (name, title, department) and ask which one. Never pick one yourself.
- "Emails from/with Dana": `outlook_search_mails_by_person(person="Dana", direction="from")`. It resolves the name first and returns candidates instead of results if it's ambiguous.

### Other people's calendars and finding a time

- "When is Ann free?" / "Find a time with Ann and Bob": use `outlook_find_meeting_times` (or `outlook_get_free_busy` for raw busy blocks). Both use Exchange free/busy, which works for anyone in the organization without calendar permissions, and include the user's own calendar by default.
- Ask or infer the working week: for a Sunday–Thursday week pass `working_days=["sun","mon","tue","wed","thu"]`. Default hours are 09:00–17:00 local.
- Offer the user a few slots, then `outlook_create_event` with the attendees (saved unsent in draft-only mode).
- "What's on Ann's calendar?" (meeting titles/details): `outlook_list_events(mailbox="ann@...")`. This needs Ann to have shared her calendar; if it fails, fall back to free/busy and say so.

## Default workflow

1. **Identify the item.** For mail, prefer `search_mails` over `list_mails` when the user describes content (subject keyword, sender name). For calendar, `list_events` with a date range. Capture `entry_id`s as you go (`response_format='json'` makes this clean).
2. **Read the full record before editing or replying.** `get_mail` / `get_event` return the full body — list/search results only include a 200-char preview, which is not enough to write a grounded reply.
3. **Confirm destructive or outbound actions** unless explicitly authorized. State the verb, recipients, and key fields. For replies, show the user the body you're about to send, or save it as a draft and tell them where to find it.
4. **Act, then report briefly.** "Sent — replied-all to 'Re: Q3 budget' with the revised numbers." Don't echo raw JSON unless asked.

## Critical gotchas

Read `references/gotchas.md` for the complete list. The non-negotiables:

- **Exchange sender addresses look like `EX:/O=...` distinguished names**, not SMTP. When filtering with `from_address` in `list_mails`, pass a name substring (e.g. `"sarah"`) rather than an exact address.
- **Mail rule toggling is live.** `toggle_rule` flips a real rule the instant it's called. Always `list_rules` first, confirm the exact rule name with the user, then toggle.
- **OOO is read-only here.** `get_out_of_office` reports state; there is no tool to enable/disable. If the user wants to set OOO, tell them to use Outlook → File → Automatic Replies.
- **Programmatic Access prompts** on corporate machines can silently block `send_mail`/`reply_mail`/`forward_mail`/`delete_mail`. If the user reports "you said you sent it but I don't see it", point them to Outlook → File → Options → Trust Center → Programmatic Access.
- **First call after a cold start can take several seconds.** Don't retry as if it had failed.
- **A meeting that has attendees is sent the moment `create_event` succeeds** unless you pass `send_invites=false`. Confirm attendee list and times before calling.
- **Recurring meetings: an occurrence's `entry_id` is the series id.** To change, delete or respond to one occurrence, pass `occurrence_start` (that occurrence's `start` exactly as `list_events` returned it). Changing or deleting the whole series requires `apply_to_series=true` — ask the user which they mean.
- **Signatures:** pass `signature=true` (the account default) or a name from `outlook_list_signatures`. Replies may already contain Outlook's automatic signature; the server won't add it twice.
- **Only classic Outlook is supported.** If the user is on the new Outlook (`olk.exe`), the tools won't see it. Tell them to switch back to classic Outlook.

## Reference files (load on demand)

- `references/tools.md` — every `outlook_*` tool: parameters, defaults, return shape, and notes on which fields feed into which subsequent tools.
- `references/recipes.md` — worked examples for triage, drafting replies, weekly digests, scheduling meetings, recurring events, attachment workflows, multi-mailbox setups.
- `references/gotchas.md` — every quirk and failure mode worth knowing, with the user-visible symptom and the corrective action.
- `references/setup.md` — load this **only when the `outlook_*` tools aren't yet available**. It walks the user through the one-time install of the underlying integration (Windows + classic Outlook prerequisites, install command, registering with the user's MCP client, and a troubleshooting tree).

Read the relevant reference file when you're about to do something for the first time in a session. The information is dense — fetching it once and keeping it in working memory is cheaper than calling tools blindly and recovering from errors.
