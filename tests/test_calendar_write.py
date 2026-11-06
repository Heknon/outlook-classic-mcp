"""Tests for calendar create/update/delete/respond against fake COM objects."""

import datetime as dt

import pytest

from outlook_mcp.client import calendar as cal
from outlook_mcp.errors import OutlookError
from tests._fakes import FakeAppt, FakeNamespace, fake_outlook_creating

UTC = dt.timezone.utc


def test_all_day_event_spans_whole_days():
    appt = FakeAppt()
    out = cal.create_event(
        fake_outlook_creating(appt), None, subject="Holiday", start="2026-12-24", all_day=True
    )
    assert appt.AllDayEvent is True
    assert appt.Start == dt.datetime(2026, 12, 24, tzinfo=UTC)  # to_com relabel
    assert appt.End == dt.datetime(2026, 12, 25, tzinfo=UTC)
    assert out["all_day"] is True


def test_all_day_inclusive_end_and_time_is_dropped():
    appt = FakeAppt()
    cal.create_event(
        fake_outlook_creating(appt), None, subject="Trip",
        start="2026-12-24T09:30:00", end="2026-12-26", all_day=True,
    )
    assert appt.Start == dt.datetime(2026, 12, 24, tzinfo=UTC)
    assert appt.End == dt.datetime(2026, 12, 26, tzinfo=UTC)


def test_timed_event_requires_end():
    with pytest.raises(OutlookError, match="end is required"):
        cal.create_event(fake_outlook_creating(FakeAppt()), None, subject="x", start="2026-01-01T10:00")


def test_meeting_can_be_saved_without_sending_invites():
    appt = FakeAppt()
    out = cal.create_event(
        fake_outlook_creating(appt), None, subject="Sync",
        start="2026-05-01T10:00", end="2026-05-01T10:30",
        attendees=["a@x.com"], optional_attendees=["b@x.com"], send_invites=False,
    )
    assert appt.events == ["Save"]
    assert out["invites_sent"] is False
    assert [(r.Address, r.Type) for r in appt.Recipients] == [("a@x.com", 1), ("b@x.com", 2)]


def test_meeting_sends_invites_by_default_but_not_in_draft_only(monkeypatch):
    appt = FakeAppt()
    cal.create_event(
        fake_outlook_creating(appt), None, subject="Sync",
        start="2026-05-01T10:00", end="2026-05-01T10:30", attendees=["a@x.com"],
    )
    assert appt.events == ["Save", "Send"]

    monkeypatch.setenv("OUTLOOK_MCP_DRAFT_ONLY", "1")
    appt2 = FakeAppt()
    out = cal.create_event(
        fake_outlook_creating(appt2), None, subject="Sync",
        start="2026-05-01T10:00", end="2026-05-01T10:30", attendees=["a@x.com"],
    )
    assert appt2.events == ["Save"]
    assert out["invites_sent"] is False


def test_update_refuses_whole_series_without_opt_in():
    series = FakeAppt(recurrence_state=1)
    ns = FakeNamespace({"S": series})
    with pytest.raises(OutlookError, match="recurring series"):
        cal.update_event(None, ns, entry_id="S", subject="New")
    cal.update_event(None, ns, entry_id="S", subject="New", apply_to_series=True)
    assert series.Subject == "New"


def test_update_single_occurrence():
    occ = FakeAppt(recurrence_state=2)
    series = FakeAppt(recurrence_state=1, occurrence=occ)
    ns = FakeNamespace({"S": series})
    cal.update_event(
        None, ns, entry_id="S", occurrence_start="2026-05-04T09:00:00",
        start="2026-05-04T11:00:00", end="2026-05-04T11:30:00",
    )
    assert series._pattern.asked == [dt.datetime(2026, 5, 4, 9, 0, tzinfo=UTC)]
    assert occ.Start == dt.datetime(2026, 5, 4, 11, 0, tzinfo=UTC)
    assert series.Start is None  # the series itself was not moved


def test_update_sends_to_attendees_only_when_asked_and_organizer():
    m = FakeAppt(meeting=True, organizer=True)
    ns = FakeNamespace({"M": m})
    out = cal.update_event(None, ns, entry_id="M", location="Room 2")
    assert m.events == ["Save"] and out["update_sent"] is False
    assert any("not notified" in n for n in out["notes"])

    out = cal.update_event(None, ns, entry_id="M", location="Room 3", send_update=True)
    assert m.events[-2:] == ["Save", "Send"] and out["update_sent"] is True

    attendee = FakeAppt(meeting=True, organizer=False)
    ns2 = FakeNamespace({"A": attendee})
    out = cal.update_event(None, ns2, entry_id="A", subject="x", send_update=True)
    assert "Send" not in attendee.events
    assert any("not the organizer" in n for n in out["notes"])


def test_update_attendees():
    m = FakeAppt(meeting=True)
    m.Recipients.Add("a@x.com")
    m.Recipients.Add("b@x.com")
    ns = FakeNamespace({"M": m})
    out = cal.update_event(
        None, ns, entry_id="M", add_attendees=["c@x.com"], remove_attendees=["A@x.com"]
    )
    assert [r.Address for r in m.Recipients] == ["b@x.com", "c@x.com"]
    assert out["removed_attendees"] == ["a@x.com"]


def test_delete_meeting_sends_cancellation_first():
    m = FakeAppt(meeting=True, organizer=True)
    ns = FakeNamespace({"M": m})
    out = cal.delete_event(None, ns, entry_id="M")
    assert m.MeetingStatus == 5
    assert m.events == ["Save", "Send", "Delete"]
    assert out["cancellation_sent"] is True


def test_delete_meeting_in_draft_only_mode_refuses_to_send(monkeypatch):
    monkeypatch.setenv("OUTLOOK_MCP_DRAFT_ONLY", "1")
    m = FakeAppt(meeting=True, organizer=True)
    ns = FakeNamespace({"M": m})
    with pytest.raises(OutlookError, match="draft-only"):
        cal.delete_event(None, ns, entry_id="M")
    assert m.events == []
    cal.delete_event(None, ns, entry_id="M", send_cancellation=False)
    assert m.events == ["Delete"]


def test_delete_as_attendee_just_deletes():
    m = FakeAppt(meeting=True, organizer=False)
    out = cal.delete_event(None, FakeNamespace({"M": m}), entry_id="M")
    assert m.events == ["Delete"]
    assert out["cancellation_sent"] is False


class _Resp:
    def __init__(self):
        self.Body = ""
        self.events = []

    def Send(self):
        self.events.append("Send")

    def Save(self):
        self.events.append("Save")


class _Invite(FakeAppt):
    def __init__(self):
        super().__init__(meeting=True, organizer=False)
        self.resp = _Resp()

    def Respond(self, code, no_ui):
        self.code = code
        return self.resp


def test_respond_with_comment_and_draft_only(monkeypatch):
    inv = _Invite()
    ns = FakeNamespace({"I": inv})
    cal.respond_event(None, ns, entry_id="I", response="decline", comment="Out that day")
    assert inv.code == 4 and inv.resp.Body == "Out that day" and inv.resp.events == ["Send"]

    monkeypatch.setenv("OUTLOOK_MCP_DRAFT_ONLY", "1")
    inv2 = _Invite()
    out = cal.respond_event(None, FakeNamespace({"I": inv2}), entry_id="I", response="accept")
    assert inv2.resp.events == ["Save"]
    assert out["status"] == "response_saved_to_drafts"
