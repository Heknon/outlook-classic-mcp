---
name: brief
description: >-
  Short Outlook briefing for the start of the day: what needs triage, open Action items, replies that arrived for Waiting items, and today's meetings with conflicts. Use when the user asks for a brief, morning summary, "what's on today", "where do I stand", or invokes /outlook:brief.
---

# Outlook brief

Keep it to one screen. Read-only: change nothing unless the user asks afterwards.

1. **Buckets** — `outlook_list_quick_steps` to find the user's Action / Waiting / Read folders (or use folder names the user gave earlier).
2. **Inbox** — `outlook_list_mails(folder="inbox", limit=50, response_format="json")`: count, unread count, and the 3 most important-looking new messages (from their manager or direct team, with questions or deadlines).
3. **Action** — list the Action folder (`outlook_list_mails(folder="<action folder>")`), oldest first: subject, from, age in days. Call out anything older than 7 days.
4. **Waiting** — `outlook_check_waiting()`: which waiting items got replies (who, one-line gist) and the longest-waiting ones still open (offer to draft a nudge).
5. **Today** — `outlook_list_events(start=<today 00:00>, end=<today 23:59>)`: meetings in order, gaps, overlapping meetings, and anything starting in the next hour. Tomorrow's first meeting if it's early.
6. **Close** with up to three suggested next steps, e.g. "triage 14 inbox messages (/outlook:triage)", "reply to Dana about Q3 (waiting 9 days)".

Use the user's local times as returned. Don't paste raw JSON.
