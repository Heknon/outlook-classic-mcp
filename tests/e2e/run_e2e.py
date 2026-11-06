"""End-to-end: the real MCP server over stdio against a simulated Outlook.

Run directly for a PASS/FAIL report:

    python tests/e2e/run_e2e.py

or via pytest (tests/e2e/test_e2e.py). The server is started exactly as
an MCP client would start it (``python -m outlook_mcp``), with stand-ins
for pywin32 on the path and TZ=Asia/Jerusalem so local-time handling is
exercised away from UTC.
"""

from __future__ import annotations

import asyncio
import datetime as dt
import json
import os
import sys
import tempfile
import traceback
from pathlib import Path

from mcp import ClientSession
from mcp.client.stdio import StdioServerParameters, stdio_client

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
TZ = "Asia/Jerusalem"


def server_params(extra_env: dict[str, str] | None = None) -> StdioServerParameters:
    env = {
        "PYTHONPATH": os.pathsep.join([str(HERE / "fakewin"), str(ROOT / "src")]),
        "TZ": TZ,
        "OUTLOOK_MCP_LOG": "WARNING",
        "PATH": os.environ.get("PATH", ""),
        "HOME": os.environ.get("HOME", tempfile.gettempdir()),
        "USERPROFILE": tempfile.gettempdir(),
        "OUTLOOK_MCP_SIGNATURES_DIR": str(HERE / "signatures"),
        **(extra_env or {}),
    }
    if sys.platform == "win32":
        env["SYSTEMROOT"] = os.environ.get("SYSTEMROOT", "C:\\Windows")
    return StdioServerParameters(command=sys.executable, args=["-m", "outlook_mcp"], env=env)


class Run:
    def __init__(self, session: ClientSession):
        self.s = session
        self.results: list[tuple[str, str, str]] = []

    async def call(self, tool: str, **args):
        res = await self.s.call_tool(tool, {**args, "response_format": "json"} if tool in JSON_FORMAT else args)
        text = "".join(getattr(c, "text", "") for c in res.content)
        if res.isError:
            raise ToolError(text)
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            return text

    async def check(self, name, coro):
        try:
            detail = await coro
            self.results.append(("PASS", name, str(detail or "")))
        except AssertionError as exc:
            self.results.append(("FAIL", name, str(exc) or "assertion failed"))
        except Exception as exc:  # noqa: BLE001
            self.results.append(("FAIL", name, f"{type(exc).__name__}: {exc}"[:300]))
            if os.environ.get("E2E_TRACE"):
                traceback.print_exc()


class ToolError(Exception):
    pass


JSON_FORMAT = {
    "outlook_list_mails", "outlook_search_mails", "outlook_search_mails_by_person",
    "outlook_get_mail", "outlook_list_events", "outlook_get_event",
}


def local(day_offset: int, hhmm: str) -> str:
    d = dt.date.today() + dt.timedelta(days=day_offset)
    return f"{d.isoformat()}T{hhmm}:00"


async def scenario_draft_only(r: Run):
    tools = {t.name for t in (await r.s.list_tools()).tools}

    async def t_tools():
        expected = {
            "outlook_reply_mail", "outlook_send_draft", "outlook_find_meeting_times",
            "outlook_get_free_busy", "outlook_search_mails_by_person", "outlook_resolve_name",
            "outlook_list_mailboxes",
        }
        assert expected <= tools, expected - tools
        return f"{len(tools)} tools"

    async def t_whoami():
        out = await r.call("outlook_whoami")
        assert "Ori Harel" in json.dumps(out), out
        return "Ori Harel"

    async def t_list():
        out = await r.call("outlook_list_mails", folder="inbox", limit=10)
        subjects = [i["subject"] for i in out["items"]]
        assert subjects[0] == "Q3 numbers" or "Early bird" in subjects, subjects
        assert out["items"][0]["received"].endswith("+02:00"), out["items"][0]["received"]
        return f"{len(subjects)} mails, local offset +02:00"

    async def t_since_midnight():
        midnight = local(0, "00:00")
        out = await r.call("outlook_list_mails", folder="inbox", since=midnight, limit=50)
        subjects = {i["subject"] for i in out["items"]}
        assert "Early bird" in subjects, f"00:30 mail missed by 'since midnight' (UTC bug): {subjects}"
        assert "Design review" not in subjects, subjects
        return f"{sorted(subjects)}"

    async def t_unread_and_since():
        out = await r.call("outlook_list_mails", folder="inbox", unread_only=True, since=local(-1, "00:00"))
        subjects = {i["subject"] for i in out["items"]}
        assert subjects == {"Early bird", "Q3 numbers"}, subjects
        return "two unread"

    async def t_search():
        out = await r.call("outlook_search_mails", query="Q3 numbers", folder="inbox")
        assert [i["subject"] for i in out["items"]] == ["Q3 numbers"], out
        return "subject+body search ok"

    async def t_ctrlk_unique():
        out = await r.call("outlook_resolve_name", name="Yossi")
        assert out["status"] == "resolved" and out["smtp_address"] == "yossi@contoso.com", out
        return "Yossi -> yossi@contoso.com"

    async def t_ctrlk_ambiguous():
        out = await r.call("outlook_resolve_name", name="Dana")
        assert out["status"] == "ambiguous", out
        got = {c["smtp_address"] for c in out["candidates"]}
        assert got == {"dana.levi@contoso.com", "dcohen@contoso.com"}, got
        return "Dana -> 2 candidates"

    async def t_ctrlk_alias():
        out = await r.call("outlook_resolve_name", name="dcohen")
        assert out["status"] == "resolved" and out["person"]["department"] == "Design", out
        return "alias lookup"

    async def t_by_person():
        out = await r.call("outlook_search_mails_by_person", person="Dana Levi", direction="from")
        subjects = sorted(i["subject"] for i in out["items"])
        assert subjects == ["Old thread", "Q3 numbers"], subjects
        return "2 from Dana Levi"

    async def t_by_person_cc():
        out = await r.call("outlook_search_mails_by_person", person="Yossi Mizrahi", direction="to")
        subjects = [i["subject"] for i in out["items"]]
        assert subjects == ["Q3 numbers"], subjects  # Yossi was CC'd
        return "found via CC"

    async def t_by_person_ambiguous():
        out = await r.call("outlook_search_mails_by_person", person="Dana")
        assert out["status"] == "ambiguous" and out["count"] == 0, out
        return "asks which Dana"

    async def t_reply_default_is_draft():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=10)
        q3 = next(i for i in mails["items"] if i["subject"] == "Q3 numbers")
        out = await r.call("outlook_reply_mail", entry_id=q3["entry_id"], body="Thanks Dana,\nI'll look today.")
        assert out["status"] == "saved_to_drafts", out
        assert any("Draft-only" in n for n in out.get("notes", [])), out
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"], include_html=True)
        html = draft["html_body"]
        assert html.index("Thanks Dana") > html.index("<div class=WordSection1>"), "text not inside body"
        assert html.index("Thanks Dana") < html.index("From:"), "text not above quote"
        assert "Q3 numbers" in draft["subject"]
        r.state["draft"] = out["entry_id"]
        return "threaded draft, text above quote"

    async def t_reply_all_extra():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=10)
        q3 = next(i for i in mails["items"] if i["subject"] == "Q3 numbers")
        out = await r.call(
            "outlook_reply_mail", entry_id=q3["entry_id"], body="+Dana Cohen", reply_all=True,
            cc=["dcohen@contoso.com"], subject="RE: Q3 numbers (updated)", importance="high",
        )
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"])
        assert "Dana Cohen" in draft["cc"] and "Yossi" in draft["cc"], draft["cc"]
        assert draft["subject"] == "RE: Q3 numbers (updated)" and draft["importance"] == 2, draft
        return f"cc: {draft['cc']}"

    async def t_plain_reply():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=10)
        note = next(i for i in mails["items"] if i["subject"] == "Plain text note")
        out = await r.call("outlook_reply_mail", entry_id=note["entry_id"], body="ok")
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"])
        assert draft["body"].startswith("ok\n\n") and "Plain text note" in draft["body"], draft["body"][:80]
        return "plain stays plain"

    async def t_named_signature():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=10)
        dr = next(i for i in mails["items"] if i["subject"] == "Design review")
        out = await r.call("outlook_forward_mail", entry_id=dr["entry_id"], to=["yossi@contoso.com"],
                           body="FYI", signature="Work")
        assert out["signature"] == "Work", out
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"], include_html=True)
        html = draft["html_body"]
        assert html.index("FYI") < html.index("Ori Harel | R&amp;D") < html.index("From:"), "order"
        assert "cid:" in html and draft["attachment_count"] == 1, draft["attachments"]
        return "signature + inline image"

    async def t_send_draft_refused():
        try:
            await r.call("outlook_send_draft", entry_id=r.state["draft"])
        except ToolError as exc:
            assert "draft-only" in str(exc).lower(), exc
            return "refused"
        raise AssertionError("send_draft was not refused in default mode")

    async def t_drafts_folder():
        out = await r.call("outlook_list_mails", folder="drafts", limit=20)
        assert out["count"] >= 4, out["count"]
        sent = await r.call("outlook_list_mails", folder="sent", limit=20)
        assert sent["count"] == 0, f"something was SENT: {[i['subject'] for i in sent['items']]}"
        return f"{out['count']} drafts, 0 sent"

    async def t_event_time():
        out = await r.call("outlook_create_event", subject="Smoke", start=local(2, "14:00"),
                           end=local(2, "14:30"), reminder_minutes=None)
        ev = await r.call("outlook_get_event", entry_id=out["entry_id"])
        assert ev["start"] == local(2, "14:00") + "+02:00", ev["start"]
        r.state["event"] = out["entry_id"]
        return ev["start"]

    async def t_all_day():
        out = await r.call("outlook_create_event", subject="Holiday", start=(dt.date.today() + dt.timedelta(days=3)).isoformat(), all_day=True)
        ev = await r.call("outlook_get_event", entry_id=out["entry_id"])
        assert ev["all_day"] and ev["start"].startswith(local(3, "00:00")) and ev["end"].startswith(local(4, "00:00")), ev
        return "midnight to midnight"

    async def t_meeting_unsent():
        out = await r.call("outlook_create_event", subject="Planning", start=local(2, "16:00"),
                           end=local(2, "16:30"), attendees=["Yossi"])
        assert out["invites_sent"] is False, out
        return "invite saved, not sent"

    async def t_list_events_window():
        out = await r.call("outlook_list_events", start=local(1, "00:00"), end=local(1, "23:59"))
        subjects = [i["subject"] for i in out["items"]]
        assert "Existing 1:1" in subjects and "Standup" in subjects, subjects
        standup = next(i for i in out["items"] if i["subject"] == "Standup")
        assert standup["recurrence"] == "occurrence", standup
        r.state["standup"] = standup
        return f"{subjects}"

    async def t_series_guard():
        st = r.state["standup"]
        try:
            await r.call("outlook_update_event", entry_id=st["entry_id"], location="Room 2")
        except ToolError as exc:
            assert "recurring series" in str(exc), exc
        else:
            raise AssertionError("series edit not guarded")
        out = await r.call("outlook_update_event", entry_id=st["entry_id"],
                           occurrence_start=st["start"], location="Room 2")
        assert out["status"] == "updated", out
        return "guarded; single occurrence updated"

    async def t_shared_calendar():
        out = await r.call("outlook_list_events", mailbox="Dana Levi", start=local(1, "00:00"), end=local(1, "23:59"))
        assert [i["subject"] for i in out["items"]] == ["Roadmap planning", "Maybe: vendor call"], out
        return "Dana Levi's shared calendar"

    async def t_shared_denied():
        try:
            await r.call("outlook_list_events", mailbox="Yossi")
        except ToolError as exc:
            assert "free/busy" in str(exc), exc
            return "points to free/busy"
        raise AssertionError("expected permission error")

    async def t_free_busy():
        out = await r.call("outlook_get_free_busy", people=["Yossi", "dana.levi@contoso.com", "ext@gmail.com"],
                           start=local(1, "08:00"), end=local(1, "18:00"))
        by = {p["name"]: p["busy"] for p in out["people"]}
        assert by["Yossi Mizrahi"] == [{"start": local(1, "11:00") + "+02:00", "end": local(1, "13:00") + "+02:00", "status": "busy"}], by
        statuses = [b["status"] for b in by["Dana Levi"]]
        assert statuses == ["busy", "tentative"], statuses
        assert out["problems"][0]["who"] == "ext@gmail.com", out.get("problems")
        return "Yossi 11-13, Dana busy+tentative, external reported"

    async def t_find_times():
        out = await r.call(
            "outlook_find_meeting_times", attendees=["Yossi", "dana.levi@contoso.com"],
            duration_minutes=60, start=local(1, "00:00"), end=local(1, "23:59"),
            working_days=["sun", "mon", "tue", "wed", "thu", "fri", "sat"], max_per_day=10,
        )
        starts = [s["start"][11:16] for s in out["slots"]]
        # Me: 09:00-09:15 standup, 10-11 1:1; Yossi 11-13; Dana 13-15, tentative 15-16.
        assert starts == ["16:00"], starts
        return f"slots {starts}"

    async def t_find_times_tentative():
        out = await r.call(
            "outlook_find_meeting_times", attendees=["Yossi", "dana.levi@contoso.com"],
            optional_attendees=["dcohen"], duration_minutes=60, start=local(1, "00:00"),
            end=local(1, "23:59"), tentative_is_free=True, max_per_day=10,
            working_days=["sun", "mon", "tue", "wed", "thu", "fri", "sat"],
        )
        starts = [s["start"][11:16] for s in out["slots"]]
        assert starts == ["15:00", "16:00"], starts
        assert out["slots"][0].get("tentative") == ["Dana Levi"], out["slots"][0]
        return f"slots {starts}"

    async def t_delete_meeting_refused():
        cal = await r.call("outlook_list_events", start=local(2, "00:00"), end=local(2, "23:59"))
        planning = next(i for i in cal["items"] if i["subject"] == "Planning")
        try:
            await r.call("outlook_delete_event", entry_id=planning["entry_id"])
        except ToolError as exc:
            assert "draft-only" in str(exc).lower()
            return "cancellation refused"
        raise AssertionError("cancellation not refused")

    async def t_attachment_sandbox():
        try:
            await r.call("outlook_send_mail", to=["yossi@contoso.com"], subject="x", body="x",
                         attachments=[str(ROOT / "README.md")])
        except ToolError as exc:
            assert "outside the user profile" in str(exc), exc
            return "blocked"
        raise AssertionError("attachment outside profile accepted")

    async def t_mailboxes():
        out = await r.call("outlook_list_mailboxes")
        kinds = {m["name"]: m["kind"] for m in out["mailboxes"]}
        assert kinds == {
            "ori.harel@contoso.com": "primary",
            "Online Archive - ori.harel@contoso.com": "archive",
            "Support Team": "shared",
            "Old PST 2019": "pst_or_other",
        }, kinds
        return "primary, archive, shared, PST"

    async def t_folder_scope_unchanged():
        out = await r.call("outlook_search_mails", query="budget")
        assert out["count"] == 0 and out["folder"] == "Inbox", out
        return "default still just Inbox"

    async def t_archive():
        out = await r.call("outlook_search_mails", query="budget", search_in="archive")
        got = [(i["subject"], i["folder"]) for i in out["items"]]
        assert got == [
            ("Budget 2024 final", "Online Archive - ori.harel@contoso.com/Inbox"),
            ("Apollo kickoff budget", "Online Archive - ori.harel@contoso.com/Inbox/Projects/Apollo"),
        ], got
        return "2 in archive incl. subfolder, deleted skipped"

    async def t_archive_deleted():
        out = await r.call("outlook_search_mails", query="budget", search_in="archive", include_deleted=True)
        assert out["count"] == 3, [i["subject"] for i in out["items"]]
        return "3 with Deleted Items"

    async def t_all():
        out = await r.call("outlook_search_mails", query="budget", search_in="all")
        subjects = [i["subject"] for i in out["items"]]
        assert subjects == ["Customer budget question", "Budget 2024 final", "Apollo kickoff budget", "Wedding budget"], subjects
        return f"{len(subjects)} across 4 mailboxes, newest first"

    async def t_named_mailbox():
        out = await r.call("outlook_search_mails", query="budget", mailboxes=["support"])
        assert [i["subject"] for i in out["items"]] == ["Customer budget question"], out
        return "Support Team only"

    async def t_multiword_all():
        out = await r.call("outlook_search_mails", query="budget numbers", search_in="all")
        assert [i["subject"] for i in out["items"]] == ["Budget 2024 final"], out
        return "AND of words across mailboxes"

    async def t_person_all():
        out = await r.call("outlook_search_mails_by_person", person="Dana Levi", direction="from", search_in="all")
        subjects = [i["subject"] for i in out["items"]]
        assert subjects == ["Q3 numbers", "Old thread", "Budget 2024 final"], subjects
        return "inbox + archive"

    async def t_bad_mailbox():
        try:
            await r.call("outlook_search_mails", query="x", mailboxes=["nope"])
        except ToolError as exc:
            assert "Available:" in str(exc), exc
            return "lists available mailboxes"
        raise AssertionError("unknown mailbox accepted")

    for name, fn in [
        ("list mailboxes", t_mailboxes), ("search default scope unchanged", t_folder_scope_unchanged),
        ("search online archive", t_archive), ("archive incl. deleted", t_archive_deleted),
        ("search all mailboxes", t_all), ("search named mailbox", t_named_mailbox),
        ("multi-word across mailboxes", t_multiword_all), ("person across mailboxes", t_person_all),
        ("unknown mailbox name", t_bad_mailbox),
    ]:
        await r.check(name, fn())

    for name, fn in [
        ("server lists the new tools", t_tools), ("whoami", t_whoami), ("list inbox", t_list),
        ("since-midnight filter in UTC+2", t_since_midnight), ("unread + since filters", t_unread_and_since),
        ("search subject/body", t_search), ("Ctrl+K unique", t_ctrlk_unique),
        ("Ctrl+K ambiguous", t_ctrlk_ambiguous), ("Ctrl+K alias", t_ctrlk_alias),
        ("mails from person", t_by_person), ("mails to person (CC)", t_by_person_cc),
        ("ambiguous person search", t_by_person_ambiguous), ("reply is a draft by default", t_reply_default_is_draft),
        ("reply-all with extras", t_reply_all_extra), ("plain-text reply", t_plain_reply),
        ("forward with named signature", t_named_signature), ("send_draft refused by default", t_send_draft_refused),
        ("nothing sent", t_drafts_folder), ("event time round-trip UTC+2", t_event_time), ("all-day event", t_all_day),
        ("meeting invite not sent", t_meeting_unsent), ("list events + recurrence", t_list_events_window),
        ("recurring series guard", t_series_guard), ("colleague's shared calendar", t_shared_calendar),
        ("unshared calendar error", t_shared_denied), ("free/busy", t_free_busy),
        ("find meeting times", t_find_times), ("find times, tentative ok + optional", t_find_times_tentative),
        ("meeting cancellation refused", t_delete_meeting_refused), ("attachment sandbox", t_attachment_sandbox),
    ]:
        await r.check(name, fn())


async def scenario_send_enabled(r: Run):
    async def t_send():
        out = await r.call("outlook_send_mail", to=["Yossi"], subject="Hello", body="Hi")
        assert out["status"] == "sent", out
        sent = await r.call("outlook_list_mails", folder="sent", limit=5)
        assert [i["subject"] for i in sent["items"]] == ["Hello"], sent
        return "sent and in Sent Items"

    async def t_unresolved():
        try:
            await r.call("outlook_send_mail", to=["Dana"], subject="x", body="x")
        except ToolError as exc:
            assert "could not be resolved" in str(exc) and "Dana" in str(exc), exc
            return "ambiguous recipient blocks send"
        raise AssertionError("sent to ambiguous 'Dana'")

    async def t_send_draft():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=10)
        q3 = next(i for i in mails["items"] if i["subject"] == "Q3 numbers")
        d = await r.call("outlook_reply_mail", entry_id=q3["entry_id"], body="draft first", save_only=True)
        assert d["status"] == "saved_to_drafts" and not d.get("notes"), d
        out = await r.call("outlook_send_draft", entry_id=d["entry_id"])
        assert out["status"] == "sent" and "Dana Levi" in out["to"], out
        return "reviewed draft sent"

    for name, fn in [("send when enabled", t_send), ("unresolved recipient", t_unresolved),
                     ("send_draft when enabled", t_send_draft)]:
        await r.check(name, fn())


async def run_all() -> list[tuple[str, str, str]]:
    results = []
    for label, env, scenario in [
        ("default (draft-only)", {}, scenario_draft_only),
        ("OUTLOOK_MCP_ALLOW_SEND=1", {"OUTLOOK_MCP_ALLOW_SEND": "1"}, scenario_send_enabled),
    ]:
        async with stdio_client(server_params(env)) as (read, write):
            async with ClientSession(read, write) as session:
                await session.initialize()
                r = Run(session)
                r.state = {}
                await scenario(r)
                results += [(s, f"[{label}] {n}", d) for s, n, d in r.results]
    return results


def main() -> int:
    results = asyncio.run(run_all())
    width = max(len(n) for _, n, _ in results)
    for status, name, detail in results:
        print(f"  {status}  {name:<{width}}  {detail}")
    failed = sum(1 for s, _, _ in results if s != "PASS")
    print(f"\n{len(results) - failed} passed, {failed} failed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
