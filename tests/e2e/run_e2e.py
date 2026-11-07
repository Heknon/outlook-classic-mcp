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
TZ = os.environ.get("E2E_TZ", "Asia/Jerusalem")


_ATT_DIR: list[str] = []


def _attachments_dir() -> str:
    """A fresh attachments folder per run (inside the e2e 'user profile')."""
    if not _ATT_DIR:
        _ATT_DIR.append(tempfile.mkdtemp(prefix="outlook-mcp-e2e-att-", dir=tempfile.gettempdir()))
    return _ATT_DIR[0]


def server_params(extra_env: dict[str, str] | None = None) -> StdioServerParameters:
    env = {
        "PYTHONPATH": os.pathsep.join([str(HERE / "fakewin"), str(ROOT / "src")]),
        "TZ": TZ,
        "OUTLOOK_MCP_LOG": "WARNING",
        "PATH": os.environ.get("PATH", ""),
        "HOME": os.environ.get("HOME", tempfile.gettempdir()),
        "USERPROFILE": tempfile.gettempdir(),
        "OUTLOOK_MCP_SIGNATURES_DIR": str(HERE / "signatures"),
        "OUTLOOK_MCP_ATTACHMENTS_DIR": _attachments_dir(),
        **(extra_env or {}),
    }
    if sys.platform == "win32":
        env["SYSTEMROOT"] = os.environ.get("SYSTEMROOT", "C:\\Windows")
    return StdioServerParameters(command=sys.executable, args=["-m", "outlook_mcp"], env=env)


_OFFSET_CACHE: dict[str, str] = {}


def server_offset(day: dt.date) -> str:
    """The UTC offset the server process will use on ``day`` ("+02:00").

    TZ=Asia/Jerusalem takes effect on Linux/macOS; Windows ignores it and
    uses the machine's zone, so ask a Python process with the same env.
    """
    key = day.isoformat()
    if key not in _OFFSET_CACHE:
        import subprocess

        code = (
            "import datetime as d;"
            f"print(d.datetime({day.year},{day.month},{day.day},12).astimezone().strftime('%z'))"
        )
        z = subprocess.run([sys.executable, "-c", code], env=server_params().env,
                           capture_output=True, text=True, check=True).stdout.strip()
        _OFFSET_CACHE[key] = f"{z[:3]}:{z[3:]}"
    return _OFFSET_CACHE[key]


_TODAY: list[dt.date] = []


def server_today() -> dt.date:
    """Today's date in the server's timezone (it differs from this
    process's date around midnight when TZ is set for the server only)."""
    if not _TODAY:
        import subprocess

        out = subprocess.run(
            [sys.executable, "-c", "import datetime as d; print(d.date.today().isoformat())"],
            env=server_params().env, capture_output=True, text=True, check=True,
        ).stdout.strip()
        _TODAY.append(dt.date.fromisoformat(out))
    return _TODAY[0]


def off(day_offset: int = 0) -> str:
    return server_offset(server_today() + dt.timedelta(days=day_offset))


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
    d = server_today() + dt.timedelta(days=day_offset)
    return f"{d.isoformat()}T{hhmm}:00"


async def scenario_draft_only(r: Run):
    tools = {t.name for t in (await r.s.list_tools()).tools}

    async def t_tools():
        expected = {
            "outlook_reply_mail", "outlook_send_draft", "outlook_find_meeting_times",
            "outlook_get_free_busy", "outlook_search_mails_by_person", "outlook_resolve_name",
            "outlook_list_mailboxes", "outlook_get_thread", "outlook_create_rule", "outlook_run_rule",
            "outlook_delete_rule", "outlook_create_category", "outlook_update_category",
            "outlook_delete_category", "outlook_list_quick_steps", "outlook_run_quick_step",
            "outlook_move_mails", "outlook_check_waiting", "outlook_calendar_stats",
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
        assert out["items"][0]["received"].endswith(off(0)), out["items"][0]["received"]
        return f"{len(subjects)} mails, local offset {off(0)}"

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
        assert subjects == ["Offsite plan", "Old thread", "Q3 numbers", "RE: Contract signature", "RE: Offsite plan"], subjects
        return "5 from Dana Levi"

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
        subjects = [i["subject"] for i in sent["items"]]
        assert subjects == ["RE: Offsite plan"], f"something was SENT: {subjects}"  # only the seeded one
        return f"{out['count']} drafts, 0 sent"

    async def t_event_time():
        out = await r.call("outlook_create_event", subject="Smoke", start=local(2, "14:00"),
                           end=local(2, "14:30"), reminder_minutes=None)
        ev = await r.call("outlook_get_event", entry_id=out["entry_id"])
        assert ev["start"] == local(2, "14:00") + off(2), ev["start"]
        r.state["event"] = out["entry_id"]
        return ev["start"]

    async def t_all_day():
        out = await r.call("outlook_create_event", subject="Holiday", start=(server_today() + dt.timedelta(days=3)).isoformat(), all_day=True)
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
        assert [i["subject"] for i in out["items"]] == ["Private appointment", "Roadmap planning", "Maybe: vendor call"], out
        assert "Katz" not in json.dumps(out) and "Clinic" not in json.dumps(out), "private details leaked"
        return "Dana Levi's shared calendar (private one masked)"

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
        assert by["Yossi Mizrahi"] == [{"start": local(1, "11:00") + off(1), "end": local(1, "13:00") + off(1), "status": "busy"}], by
        statuses = [b["status"] for b in by["Dana Levi"]]
        assert statuses == ["busy", "busy", "tentative"], statuses
        rows = {p["name"]: p for p in out["people"]}
        assert rows["Yossi Mizrahi"]["access"] == "free_busy_only" and "meetings" not in rows["Yossi Mizrahi"]
        assert rows["Dana Levi"]["access"] == "calendar", rows["Dana Levi"]
        titles = [m["subject"] for m in rows["Dana Levi"]["meetings"]]
        assert titles == ["Private appointment", "Roadmap planning", "Maybe: vendor call"], titles
        assert "Katz" not in json.dumps(out) and "Clinic" not in json.dumps(out), "private details leaked"
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
        cals = {c["name"]: c for c in out["calendars"]}
        assert cals["Dana Levi"]["access"] == "calendar" and cals["Yossi Mizrahi"]["access"] == "free_busy_only"
        mine = [m["subject"] for m in cals["Ori Harel"]["meetings"]]
        assert "Existing 1:1" in mine and "Standup" in mine, mine
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
        assert subjects == ["Customer budget question", "Budget approval", "Budget 2024 final",
                            "Apollo kickoff budget", "Wedding budget"], subjects
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
        assert subjects == ["Q3 numbers", "RE: Offsite plan", "Offsite plan", "RE: Contract signature",
                            "Old thread", "Budget 2024 final"], subjects
        return "inbox + archive"

    async def t_bad_mailbox():
        try:
            await r.call("outlook_search_mails", query="x", mailboxes=["nope"])
        except ToolError as exc:
            assert "Available:" in str(exc), exc
            return "lists available mailboxes"
        raise AssertionError("unknown mailbox accepted")

    async def t_thread():
        mails = await r.call("outlook_search_mails", query="Offsite plan", scope="subject", folder="inbox")
        last = next(i for i in mails["items"] if i["subject"] == "RE: Offsite plan")
        out = await r.call("outlook_get_thread", entry_id=last["entry_id"])
        got = [(m["from"], m["folder"], m["body"]) for m in out["messages"]]
        assert got == [
            ("Dana Levi", "Inbox", "Can we do the offsite on Thursday?"),
            ("Ori Harel", "Sent Items", "Thursday is fine, booking Room Galil.\n\nOri"),
            ("Dana Levi", "Inbox", "מעולה, תודה!"),
        ], got
        assert out["source"] == "conversation" and out["topic"] == "Offsite plan"
        return "3 messages incl. Sent Items, quotes stripped (English + Hebrew headers)"

    async def t_thread_full():
        mails = await r.call("outlook_search_mails", query="Offsite plan", scope="subject", folder="inbox")
        last = next(i for i in mails["items"] if i["subject"] == "RE: Offsite plan")
        out = await r.call("outlook_get_thread", entry_id=last["entry_id"], strip_quoted=False)
        assert "נשלח: יום חמישי" in out["messages"][2]["body"]
        return "full bodies on request"

    async def t_hebrew_reply():
        mails = await r.call("outlook_list_mails", folder="inbox", limit=20)
        note = next(i for i in mails["items"] if i["subject"] == "Plain text note")
        out = await r.call("outlook_reply_mail", entry_id=note["entry_id"], body="תודה יוסי,\nאחזור אליך מחר.")
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"], include_html=True)
        html = draft["html_body"]
        assert "dir=RTL" in html and "<span lang=HE>תודה יוסי,</span>" in html, html[:300]
        assert html.index("תודה יוסי") < html.index("Plain text note")
        return "plain-text thread answered as RTL HTML"

    async def t_hebrew_new_mail():
        out = await r.call("outlook_send_mail", to=["Yossi"], subject="פגישה מחר", body="היי יוסי,\nנפגש ב-10?")
        draft = await r.call("outlook_get_mail", entry_id=out["entry_id"], include_html=True)
        assert draft["html_body"].count("dir=RTL") == 2, draft["html_body"][:300]
        return "new Hebrew mail is RTL"

    async def t_category_crud():
        out = await r.call("outlook_create_category", name="Customer A", color="dark green")
        assert out["color"] == "dark green", out
        out = await r.call("outlook_update_category", name="customer a", color="gold")
        assert out["color"] == "dark yellow", out
        lst = await r.call("outlook_list_categories", response_format="json")
        assert {"name": "Customer A", "color": "dark yellow", "color_code": 19} in lst["items"], lst
        return "create, recolour (gold = dark yellow), list"

    async def t_categorize_mails_bulk():
        found = await r.call("outlook_search_mails", query="Invoice", scope="subject")
        ids = [i["entry_id"] for i in found["items"]]
        out = await r.call("outlook_set_category", entry_ids=ids, categories=["Finance"], mode="add", color="orange")
        assert out["count"] == 2 and out["created_categories"] == ["Finance"], out
        got = await r.call("outlook_get_mail", entry_id=ids[0])
        assert got["categories"] == "Finance", got["categories"]
        return "2 invoices tagged, 'Finance' created in orange"

    async def t_categorize_meeting():
        out = await r.call("outlook_create_event", subject="Customer A sync", start=local(3, "11:00"),
                           end=local(3, "11:30"), categories=["Customer A"])
        ev = await r.call("outlook_get_event", entry_id=out["entry_id"])
        assert ev["categories"] == "Customer A", ev
        cal = await r.call("outlook_list_events", start=local(1, "00:00"), end=local(1, "23:59"))
        st = next(i for i in cal["items"] if i["subject"] == "Standup")
        await r.call("outlook_set_category", entry_id=st["entry_id"], occurrence_start=st["start"],
                     categories="Work")
        series = await r.call("outlook_get_event", entry_id=st["entry_id"])
        assert series["categories"] != "Work", "the whole series was recoloured"
        return "new meeting coloured; one standup occurrence only"

    async def t_rule_preview_and_create():
        args = dict(name="Acme invoices",
                    conditions={"sender_address_words": ["@acme-supplies.com"], "subject_words": ["invoice"]},
                    actions={"move_to_folder": "Inbox/Vendors", "categories": ["Vendors"], "category_color": "teal"})
        pre = await r.call("outlook_create_rule", **args)
        assert pre["status"] == "preview" and pre["matches"]["would_match"] == 2, pre
        listed = await r.call("outlook_list_rules", details=True)
        assert listed["count"] == 0, "preview saved a rule"
        out = await r.call("outlook_create_rule", **args, confirm=True)
        assert out["status"] == "created", out
        acts = out["rule"]["actions"]
        assert acts["categories"] == ["Vendors"] and acts["move_to_folder"].endswith("Inbox\\Vendors"), acts
        return "previewed 2 matches, then created"

    async def t_rule_run():
        pre = await r.call("outlook_run_rule", name="acme invoices")
        assert pre["status"] == "preview" and pre["matches"]["would_match"] == 2, pre
        inbox = await r.call("outlook_search_mails", query="Invoice", scope="subject")
        assert inbox["count"] == 2, "preview moved mail"
        await r.call("outlook_run_rule", name="acme invoices", confirm=True)
        inbox = await r.call("outlook_search_mails", query="Invoice", scope="subject")
        moved = await r.call("outlook_search_mails", query="Invoice", scope="subject", folder="Inbox/Vendors")
        assert inbox["count"] == 0 and moved["count"] == 2, (inbox["count"], moved["count"])
        cats = await r.call("outlook_get_mail", entry_id=moved["items"][0]["entry_id"])
        assert "Vendors" in cats["categories"], cats["categories"]
        return "2 invoices moved to Inbox/Vendors and tagged"

    async def t_rule_forward_refused():
        try:
            await r.call("outlook_create_rule", name="fw", conditions={"subject_words": ["x"]},
                         actions={"forward_to": ["dana.levi@contoso.com"]})
        except ToolError as exc:
            assert "draft-only" in str(exc), exc
            return "refused in draft-only mode"
        raise AssertionError("forward rule accepted in draft-only mode")

    async def t_rule_delete():
        pre = await r.call("outlook_delete_rule", name="Acme invoices")
        assert pre["status"] == "preview"
        await r.call("outlook_delete_rule", name="Acme invoices", confirm=True)
        listed = await r.call("outlook_list_rules", details=True)
        assert listed["count"] == 0, listed
        return "preview, then deleted"

    for name, fn in [
        ("category create/recolour/list", t_category_crud), ("categorize mails in bulk", t_categorize_mails_bulk),
        ("colour a meeting / one occurrence", t_categorize_meeting),
        ("rule: preview then create", t_rule_preview_and_create), ("rule: run on existing mail", t_rule_run),
        ("rule: forward refused (draft-only)", t_rule_forward_refused), ("rule: delete", t_rule_delete),
    ]:
        await r.check(name, fn())

    async def t_attachment_details():
        found = await r.call("outlook_search_mails", query="Contract draft", scope="subject")
        m = await r.call("outlook_get_mail", entry_id=found["items"][0]["entry_id"])
        got = [(a["filename"], a["kind"], a["inline"]) for a in m["attachments"]]
        assert got == [("Contract v3.pdf", "pdf", False), ("image001.png", "image", True)], got
        assert m["file_attachment_count"] == 1
        assert m["linked_files"] == ["https://contoso.sharepoint.com/:w:/r/sites/Legal/Terms.docx"], m.get("linked_files")
        r.state["contract"] = found["items"][0]["entry_id"]
        return "PDF + inline logo + SharePoint link told apart"

    async def t_save_and_reply_with_it():
        out = await r.call("outlook_save_attachments", entry_id=r.state["contract"])
        assert [f["filename"] for f in out["files"]] == ["Contract v3.pdf"], out
        path = out["files"][0]["path"]
        assert os.path.isfile(path) and open(path, "rb").read().startswith(b"%PDF"), path
        assert "Contract draft" in out["output_dir"] and out["skipped"][0]["reason"] == "inline image"
        reply = await r.call("outlook_reply_mail", entry_id=r.state["contract"],
                             body="Reviewed, comments inline.", attachments=[path])
        draft = await r.call("outlook_get_mail", entry_id=reply["entry_id"])
        assert [a["filename"] for a in draft["attachments"]] == ["Contract v3.pdf"], draft["attachments"]
        return f"saved to {os.path.basename(out['output_dir'])}/, re-attached to a reply draft"

    for name, fn in [
        ("attachment details", t_attachment_details), ("save attachment, reply with it", t_save_and_reply_with_it),
    ]:
        await r.check(name, fn())

    async def t_calendar_stats():
        out = await r.call("outlook_calendar_stats", mailbox="Dana Levi")
        assert out["access"] == "calendar", out
        rec = {x["title"]: x for x in out["recurring_meetings"]}
        assert rec["Product standup"]["usual_time"] == "09:30" and rec["Product standup"]["occurrences"] >= 38, rec
        assert rec["1:1 Dana / Yossi"]["usual_days"] == ["mon"] and rec["Roadmap review"]["usual_days"] == ["wed"]
        assert out["top_one_on_ones"][0]["name"] == "Yossi Mizrahi", out["top_one_on_ones"]
        assert out["private_appointments"] == 8 and "Pilates" not in json.dumps(out), "private leaked"
        days = {d["day"] for d in out["by_weekday"]}
        assert days == {"sun", "mon", "tue", "wed", "thu"}, days
        assert out["typical_day"]["first_meeting_starts"] in ("07:30", "09:30"), out["typical_day"]
        return f"{out['meetings']} meetings/8 wks; standup 09:30 daily, 1:1 Mon, roadmap Wed; private counted not named"

    async def t_history_paging():
        rng = dict(mailbox="Dana Levi", start=local(-56, "00:00"), end=local(0, "00:00"), compact=True)
        first = await r.call("outlook_list_events", limit=25, **rng)
        assert first["truncated"] and first["next_offset"] == 25, (first["count"], first["truncated"])
        seen = [i["start"] + i["subject"] for i in first["items"]]
        offset = first["next_offset"]
        while offset is not None:
            page = await r.call("outlook_list_events", limit=25, offset=offset, **rng)
            seen += [i["start"] + i["subject"] for i in page["items"]]
            offset = page["next_offset"]
        assert len(seen) == len(set(seen)), "duplicates across pages"
        stats = await r.call("outlook_calendar_stats", mailbox="Dana Levi",
                             start=local(-56, "00:00"), end=local(0, "00:00"))
        assert len(seen) == stats["meetings"] + stats["all_day_events"], (len(seen), stats["meetings"])
        assert "Pilates" not in json.dumps(seen)
        with_people = await r.call("outlook_list_events", limit=3, include_attendees=True, **rng)
        assert any(i.get("attendees") for i in with_people["items"]), with_people["items"]
        return f"{len(seen)} past events over pages, no duplicates, matches stats"

    async def t_stats_free_busy_only():
        out = await r.call("outlook_calendar_stats", mailbox="Yossi", start=local(1, "00:00"), end=local(3, "00:00"))
        assert out["access"] == "free_busy_only" and out["meetings"] >= 1 and not out["top_people"], out
        return "unshared calendar: times only, with a note"

    for name, fn in [
        ("calendar patterns (לו״ז intelligence)", t_calendar_stats), ("calendar history paging", t_history_paging),
        ("stats without calendar access", t_stats_free_busy_only),
    ]:
        await r.check(name, fn())

    for name, fn in [
        ("whole thread in one call", t_thread), ("thread with full bodies", t_thread_full),
        ("Hebrew reply is right-to-left", t_hebrew_reply), ("Hebrew new mail is right-to-left", t_hebrew_new_mail),
    ]:
        await r.check(name, fn())

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
("since-midnight filter (local time)", t_since_midnight), ("unread + since filters", t_unread_and_since),
        ("search subject/body", t_search), ("Ctrl+K unique", t_ctrlk_unique),
        ("Ctrl+K ambiguous", t_ctrlk_ambiguous), ("Ctrl+K alias", t_ctrlk_alias),
        ("mails from person", t_by_person), ("mails to person (CC)", t_by_person_cc),
        ("ambiguous person search", t_by_person_ambiguous), ("reply is a draft by default", t_reply_default_is_draft),
        ("reply-all with extras", t_reply_all_extra), ("plain-text reply", t_plain_reply),
        ("forward with named signature", t_named_signature), ("send_draft refused by default", t_send_draft_refused),
        ("nothing sent", t_drafts_folder), ("event time round-trip (local time)", t_event_time), ("all-day event", t_all_day),
        ("meeting invite not sent", t_meeting_unsent), ("list events + recurrence", t_list_events_window),
        ("recurring series guard", t_series_guard), ("colleague's shared calendar", t_shared_calendar),
        ("unshared calendar error", t_shared_denied), ("free/busy", t_free_busy),
        ("find meeting times", t_find_times), ("find times, tentative ok + optional", t_find_times_tentative),
        ("meeting cancellation refused", t_delete_meeting_refused), ("attachment sandbox", t_attachment_sandbox),
    ]:
        await r.check(name, fn())


    # Triage last: it files Inbox mail that earlier checks look for.
    async def t_quick_steps_listed():
        out = await r.call("outlook_list_quick_steps")
        by = {q["name"]: q for q in out["quick_steps"]}
        assert set(by) == {"Archive", "Action", "Waiting", "Read", "Reply and Delete"}, set(by)
        assert by["Waiting"]["runnable"] and [a["kind"] for a in by["Waiting"]["actions"]] == ["move", "mark_read"]
        assert by["Waiting"]["actions"][0]["folder"].endswith("Inbox\\2-Waiting"), by["Waiting"]
        assert not by["Reply and Delete"]["runnable"], by["Reply and Delete"]
        return "5 read from the hidden folder (UTF-16 XML); Reply and Delete not runnable"

    async def t_check_waiting():
        out = await r.call("outlook_check_waiting")
        assert out["folder"] == "2-Waiting" and out["folder_found_by"] == "quick_step", out
        assert [a["subject"] for a in out["answered"]] == ["Contract signature"], out["answered"]
        assert out["answered"][0]["replies"][0]["preview"] == "Signed and attached.", out["answered"][0]
        assert [w["subject"] for w in out["still_waiting"]] == ["Budget approval"]
        assert out["still_waiting"][0]["days_waiting"] == 9
        return "Contract answered by Dana; Budget waiting 9 days"

    async def t_run_quick_step():
        inbox = await r.call("outlook_list_mails", folder="inbox", limit=50)
        design = next(i for i in inbox["items"] if i["subject"] == "Design review")
        out = await r.call("outlook_run_quick_step", name="archive", entry_ids=[design["entry_id"]])
        assert out["status"] == "done" and out["count"] == 1, out
        arch = await r.call("outlook_list_mails", folder="Archive")
        assert [i["subject"] for i in arch["items"]] == ["Design review"] and not arch["items"][0]["unread"]
        try:
            await r.call("outlook_run_quick_step", name="Reply and Delete", entry_ids=[design["entry_id"]])
        except ToolError as exc:
            assert "aren't run" in str(exc), exc
        else:
            raise AssertionError("Reply and Delete was run")
        return "Archive moved + marked read; Reply and Delete refused"

    async def t_move_mails():
        inbox = await r.call("outlook_list_mails", folder="inbox", limit=50)
        ids = [i["entry_id"] for i in inbox["items"] if i["subject"] in ("Plain text note", "Old thread")]
        out = await r.call("outlook_move_mails", entry_ids=ids, target_folder="Inbox/3-Read", add_categories=["Later"])
        assert out["count"] == 2, out
        rd = await r.call("outlook_list_mails", folder="Inbox/3-Read")
        assert sorted(i["subject"] for i in rd["items"]) == ["Old thread", "Plain text note"]
        return "2 filed to Read with a category"

    for name, fn in [
        ("Quick Steps read from mailbox", t_quick_steps_listed), ("Waiting: who replied", t_check_waiting),
        ("run a Quick Step", t_run_quick_step), ("bulk move", t_move_mails),
    ]:
        await r.check(name, fn())

async def scenario_send_enabled(r: Run):
    async def t_send():
        out = await r.call("outlook_send_mail", to=["Yossi"], subject="Hello", body="Hi")
        assert out["status"] == "sent", out
        sent = await r.call("outlook_list_mails", folder="sent", limit=5)
        assert [i["subject"] for i in sent["items"]] == ["Hello", "RE: Offsite plan"], sent
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

    async def t_forward_external_refused():
        try:
            await r.call("outlook_create_rule", name="fw", conditions={"subject_words": ["x"]},
                         actions={"redirect_to": ["someone@gmail.com"]})
        except ToolError as exc:
            assert "own organization" in str(exc), exc
            return "external redirect refused"
        raise AssertionError("external forward accepted")

    async def t_forward_internal_ok():
        out = await r.call("outlook_create_rule", name="fw", conditions={"subject_words": ["urgent"]},
                           actions={"forward_to": ["Yossi"]}, confirm=True)
        assert out["status"] == "created" and out["rule"]["actions"]["forward_to"] == ["yossi@contoso.com"], out
        return "internal forward allowed with sending on"

    for name, fn in [("rule: external forward refused", t_forward_external_refused),
                     ("rule: internal forward allowed", t_forward_internal_ok)]:
        await r.check(name, fn())

    for name, fn in [("send when enabled", t_send), ("unresolved recipient", t_unresolved),
                     ("send_draft when enabled", t_send_draft)]:
        await r.check(name, fn())


async def run_all() -> list[tuple[str, str, str]]:
    _TODAY.clear()  # recompute per run
    _ATT_DIR.clear()
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
