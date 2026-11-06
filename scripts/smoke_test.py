"""Smoke test against your real classic Outlook (Windows only).

Run from the repository root, with Outlook open:

    .venv\\Scripts\\python scripts\\smoke_test.py            # read checks + drafts, cleans up
    .venv\\Scripts\\python scripts\\smoke_test.py --keep     # leave the test draft/event for inspection
    .venv\\Scripts\\python scripts\\smoke_test.py --with ann@corp.com   # also test free/busy with a colleague

Safety:
* Draft-only mode is forced on for the whole run, so nothing can be sent.
* The only things created are one reply draft (to your newest inbox mail)
  and one calendar appointment tomorrow with no attendees. Both are
  deleted at the end unless --keep is given.

Prints PASS / FAIL / SKIP per check and exits non-zero if anything failed.
"""

from __future__ import annotations

import argparse
import datetime as dt
import os
import sys
import traceback
from pathlib import Path

os.environ["OUTLOOK_MCP_DRAFT_ONLY"] = "1"  # before importing the package
os.environ.pop("OUTLOOK_MCP_ALLOW_SEND", None)
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

if sys.platform != "win32":
    sys.exit("This smoke test needs Windows with classic Outlook.")

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

from outlook_mcp.client import account, calendar, mail  # noqa: E402
from outlook_mcp.client.folders import get_item_by_id  # noqa: E402

MARKER = "MCP smoke test - please ignore"
results: list[tuple[str, str, str]] = []
cleanup: list = []


def check(name: str):
    def deco(fn):
        def run(*a, **kw):
            try:
                detail = fn(*a, **kw) or ""
                status = "SKIP" if str(detail).startswith("SKIP") else "PASS"
                results.append((status, name, str(detail).removeprefix("SKIP").strip(": ")))
            except AssertionError as exc:
                results.append(("FAIL", name, str(exc)))
            except Exception as exc:  # noqa: BLE001
                results.append(("FAIL", name, f"{type(exc).__name__}: {exc}"))
                if os.environ.get("SMOKE_TRACE"):
                    traceback.print_exc()
        return run
    return deco


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keep", action="store_true", help="don't delete the test draft/event")
    ap.add_argument("--with", dest="colleague", help="a colleague's address for free/busy checks")
    args = ap.parse_args()

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    state: dict = {}

    def my_address() -> str:
        me = ns.CurrentUser
        try:
            return me.AddressEntry.GetExchangeUser().PrimarySmtpAddress
        except Exception:  # noqa: BLE001 - not Exchange, or no GAL entry
            return me.Address or me.Name

    @check("connected to mailbox")
    def t_whoami():
        info = account.whoami(outlook, ns)
        assert info.get("current_user"), "no current user"
        state["me"] = info["current_user"]
        return f"{info['current_user']}"

    @check("list inbox")
    def t_list():
        out = mail.list_mails(outlook, ns, folder="inbox", limit=5)
        assert out["count"] > 0, "inbox looks empty"
        state["newest"] = out["items"][0]
        return f"{out['count']} newest, top: {out['items'][0]['subject'][:50]!r}"

    @check("'since' date filter matches received times")
    def t_since():
        midnight = dt.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        out = mail.list_mails(outlook, ns, folder="inbox", limit=200, since=midnight.isoformat())
        # Ground truth: walk the newest items and compare times ourselves.
        items = ns.GetDefaultFolder(6).Items
        items.Sort("[ReceivedTime]", True)
        truth = 0
        for it in items:
            rt = getattr(it, "ReceivedTime", None)
            if rt is None:
                continue
            local = dt.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)
            if local < midnight:
                break
            if getattr(it, "Class", None) in (43, 53):
                truth += 1
            if truth > 200:
                break
        assert out["count"] == min(truth, 200), f"filter returned {out['count']}, expected {truth}"
        return f"{truth} mails since midnight"

    @check("reply saved as threaded draft")
    def t_reply_draft():
        newest = state.get("newest")
        if not newest:
            return "SKIP: no inbox mail"
        out = mail.reply_mail(outlook, ns, entry_id=newest["entry_id"], body=MARKER)
        assert out["status"] == "saved_to_drafts", out
        draft = get_item_by_id(ns, out["entry_id"])
        cleanup.append(draft)
        original = get_item_by_id(ns, newest["entry_id"])
        assert draft.ConversationID == original.ConversationID, "draft is not in the same conversation"
        assert draft.Sent is False
        if draft.BodyFormat == 2:
            html = draft.HTMLBody
            body_at = html.lower().find("<body")
            assert 0 <= body_at < html.find(MARKER), "reply text is not inside <body>"
        return f"draft {out['subject'][:50]!r} in Drafts, same conversation"

    @check("whole thread of the newest mail")
    def t_thread():
        newest = state.get("newest")
        if not newest:
            return "SKIP: no inbox mail"
        out = mail.get_thread(outlook, ns, entry_id=newest["entry_id"])
        assert out["count"] >= 1, out
        folders = sorted({m["folder"] for m in out["messages"] if m.get("folder")})
        return f"{out['total_in_thread']} messages via {out['source']} in {folders}"

    @check("Hebrew draft is right-to-left")
    def t_hebrew():
        newest = state.get("newest")
        if not newest:
            return "SKIP: no inbox mail"
        out = mail.reply_mail(outlook, ns, entry_id=newest["entry_id"], body="בדיקה - " + MARKER)
        draft = get_item_by_id(ns, out["entry_id"])
        cleanup.append(draft)
        assert draft.BodyFormat == 2 and "dir=RTL" in draft.HTMLBody, "not RTL HTML"
        return "dir=RTL paragraph in the draft (open it in Outlook to eyeball it)"

    @check("send_draft refused in draft-only mode")
    def t_send_refused():
        if not cleanup:
            return "SKIP: no draft"
        try:
            mail.send_draft(outlook, ns, entry_id=cleanup[0].EntryID)
        except Exception as exc:  # noqa: BLE001
            assert "draft-only" in str(exc).lower(), exc
            return "refused"
        raise AssertionError("send_draft did not refuse!")

    @check("calendar event keeps its time")
    def t_event():
        day = (dt.date.today() + dt.timedelta(days=1)).isoformat()
        out = calendar.create_event(
            outlook, ns, subject=MARKER, start=f"{day}T14:00:00", end=f"{day}T14:30:00",
            reminder_minutes=None,
        )
        ev = get_item_by_id(ns, out["entry_id"])
        cleanup.append(ev)
        got = calendar.get_event(outlook, ns, entry_id=out["entry_id"])
        assert got["start"].startswith(f"{day}T14:00"), f"stored as {got['start']}"
        assert got["end"].startswith(f"{day}T14:30"), f"stored as {got['end']}"
        return f"14:00-14:30 stored as {got['start'][11:16]}-{got['end'][11:16]}"

    @check("event shows up in list_events")
    def t_list_events():
        day = dt.date.today() + dt.timedelta(days=1)
        out = calendar.list_events(
            outlook, ns, start=f"{day}T00:00:00", end=f"{day}T23:59:00", limit=200
        )
        assert any(i["subject"] == MARKER for i in out["items"]), "test event not listed"
        return f"{out['count']} events tomorrow"

    @check("own free/busy shows the test event")
    def t_fb_me():
        day = dt.date.today() + dt.timedelta(days=1)
        out = calendar.get_free_busy(
            outlook, ns, people=[my_address()], start=f"{day}T13:00:00",
            end=f"{day}T16:00:00", interval_minutes=15,
        )
        if out.get("problems"):
            return f"SKIP: {out['problems'][0]['error']}"
        busy = out["people"][0]["busy"]
        # Free/busy is published by Exchange; a brand-new event can take a
        # minute to appear, so only report.
        hit = any(b["start"][11:16] <= "14:00" < b["end"][11:16] for b in busy)
        return "busy at 14:00" if hit else "SKIP: not published yet (normal right after creating)"

    @check("find meeting times (you only)")
    def t_find_me():
        out = calendar.find_meeting_times(
            outlook, ns, attendees=[my_address()], include_me=False,
            duration_minutes=30, max_results=3,
        )
        assert "slots" in out
        return f"{out['count']} slots, first {out['slots'][0]['start'] if out['slots'] else 'none'}"

    @check("free/busy and find times with colleague")
    def t_colleague():
        if not args.colleague:
            return "SKIP: pass --with <address>"
        fb = calendar.get_free_busy(outlook, ns, people=[args.colleague])
        assert not fb.get("problems"), fb.get("problems")
        found = calendar.find_meeting_times(outlook, ns, attendees=[args.colleague], max_results=3)
        return f"{len(fb['people'][0]['busy'])} busy blocks today; {found['count']} shared slots"

    @check("colleague's calendar (needs sharing)")
    def t_shared():
        if not args.colleague:
            return "SKIP: pass --with <address>"
        try:
            out = calendar.list_events(outlook, ns, mailbox=args.colleague, limit=5)
        except Exception as exc:  # noqa: BLE001
            if "free/busy access" in str(exc):
                return "SKIP: not shared with you (expected for most people)"
            raise
        return f"{out['count']} events visible"

    @check("Ctrl+K name lookup (your own first name)")
    def t_ctrlk():
        from outlook_mcp.client import contacts

        first = str(state.get("me", "")).split()[0] if state.get("me") else ""
        if not first:
            return "SKIP: no name"
        out = contacts.resolve_name(outlook, ns, name=first)
        mine = my_address().lower()
        if out["status"] == "resolved":
            return f"resolved to {out['smtp_address']}"
        assert out["status"] == "ambiguous", out
        addrs = [str(c.get("smtp_address", "")).lower() for c in out["candidates"]]
        return f"{len(addrs)} candidates" + (", including you" if mine in addrs else "")

    @check("mails from the newest sender, by name")
    def t_by_person():
        newest = state.get("newest")
        if not newest or not newest.get("from"):
            return "SKIP: no inbox mail"
        out = mail.search_mails_by_person(outlook, ns, person=newest["from"], direction="from", limit=5)
        if out.get("status") == "ambiguous":
            return f"SKIP: '{newest['from']}' is ambiguous ({len(out['candidates'])} people)"
        if out.get("status") == "not_found":
            return f"SKIP: '{newest['from']}' not in the address book (external sender?)"
        assert out["count"] >= 1, f"no mails found from {newest['from']}"
        return f"{out['count']} from {out['person'].get('smtp_address')}"

    @check("mailboxes (archive, shared, PSTs)")
    def t_mailboxes():
        from outlook_mcp.client import folders

        out = folders.list_mailboxes(outlook, ns)
        state["mailboxes"] = out["mailboxes"]
        return ", ".join(f"{m['name']} [{m['kind']}]" for m in out["mailboxes"])

    @check("search all mailboxes for the newest subject")
    def t_search_all():
        newest = state.get("newest")
        if not newest:
            return "SKIP: no inbox mail"
        word = max(newest["subject"].split() or ["a"], key=len)
        out = mail.search_mails(outlook, ns, query=word, search_in="all", scope="subject", limit=10)
        assert out["count"] >= 1, f"'{word}' not found anywhere"
        errs = out.get("folder_errors")
        return f"{out['count']} hits in {out['folders_searched']} folders" + (f"; {len(errs)} folder errors" if errs else "")

    @check("search the archive")
    def t_search_archive():
        if not any(m["kind"] == "archive" for m in state.get("mailboxes", [])):
            return "SKIP: no archive mailbox open in this profile"
        out = mail.search_mails(outlook, ns, query="a", search_in="archive", scope="subject", limit=3)
        return f"{out['count']} hits in {out['folders_searched']} archive folders"

    @check("signatures")
    def t_sigs():
        out = mail.list_signatures_info(outlook, ns)
        return f"{len(out['signatures'])} found; default new={out['default_new']!r}, reply={out['default_reply_forward']!r}"

    for t in (t_whoami, t_list, t_since, t_reply_draft, t_thread, t_hebrew, t_send_refused, t_event,
              t_list_events, t_fb_me, t_find_me, t_colleague, t_shared, t_ctrlk,
              t_by_person, t_mailboxes, t_search_all, t_search_archive, t_sigs):
        t()

    if not args.keep:
        for item in cleanup:
            try:
                item.Delete()
            except Exception:  # noqa: BLE001
                pass

    width = max(len(n) for _, n, _ in results)
    print()
    for status, name, detail in results:
        print(f"  {status:4}  {name:<{width}}  {detail}")
    failed = sum(1 for s, _, _ in results if s == "FAIL")
    print(f"\n{len(results) - failed} ok, {failed} failed" + ("" if args.keep else " (test items deleted)"))
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
