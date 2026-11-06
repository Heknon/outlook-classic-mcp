"""Tests for utils.scheduling and the free/busy client functions."""

import datetime as dt

import pytest

from outlook_mcp.client import calendar as cal
from outlook_mcp.errors import OutlookError
from outlook_mcp.utils.scheduling import (
    busy_blocks,
    find_slots,
    parse_free_busy,
    parse_hhmm,
    parse_working_days,
)

MON = dt.datetime(2026, 11, 9)  # a Monday, midnight


def fb_from(busy: list[tuple[str, str, str]], interval=30, days=7):
    """Build a free/busy string: busy = [(day_offset+HH:MM, HH:MM, code)]."""
    n = days * 24 * 60 // interval
    codes = ["0"] * n
    for start, end, code in busy:
        d, hm = start.split("+")
        s = int(d) * 1440 + int(hm[:2]) * 60 + int(hm[3:])
        e = int(d) * 1440 + int(end[:2]) * 60 + int(end[3:])
        for i in range(s // interval, e // interval):
            codes[i] = code
    return parse_free_busy("".join(codes), MON, interval)


def at(day, hhmm):
    h, m = map(int, hhmm.split(":"))
    return MON + dt.timedelta(days=day, hours=h, minutes=m)


def test_status_and_blocks():
    fb = fb_from([("0+09:00", "10:00", "2"), ("0+10:00", "10:30", "1"), ("0+12:00", "13:00", "3")])
    assert fb.status_at(at(0, "09:15")) == 2
    assert fb.status_at(at(0, "08:59")) == 0
    assert fb.status_at(MON - dt.timedelta(minutes=1)) == -1  # before the data
    blocks = busy_blocks(fb, at(0, "08:00"), at(0, "18:00"))
    assert [(b["start"], b["end"], b["status"]) for b in blocks] == [
        (at(0, "09:00"), at(0, "10:00"), "busy"),
        (at(0, "10:00"), at(0, "10:30"), "tentative"),
        (at(0, "12:00"), at(0, "13:00"), "out_of_office"),
    ]


def test_blocks_are_clipped_to_range():
    fb = fb_from([("0+09:00", "11:00", "2")])
    [b] = busy_blocks(fb, at(0, "09:45"), at(0, "10:15"))
    assert (b["start"], b["end"]) == (at(0, "09:45"), at(0, "10:15"))


def test_find_slots_avoids_everyones_busy_time():
    ann = fb_from([("0+09:00", "11:00", "2")])
    bob = fb_from([("0+11:00", "12:00", "2")])
    slots = find_slots(
        {"Ann": ann, "Bob": bob}, {},
        window_start=at(0, "08:00"), window_end=at(0, "18:00"),
        duration=dt.timedelta(minutes=60), max_per_day=10,
    )
    assert [s.start for s in slots] == [at(0, "12:00"), at(0, "13:00"), at(0, "14:00"), at(0, "15:00"), at(0, "16:00")]
    assert all(s.end <= at(0, "17:00") for s in slots)


def test_working_days_and_hours_respected():
    free = fb_from([])
    slots = find_slots(
        {"Ann": free}, {},
        window_start=at(4, "16:00"),  # Friday 16:00
        window_end=at(7, "00:00") + dt.timedelta(days=0) - dt.timedelta(minutes=1),
        duration=dt.timedelta(minutes=60),
        working_days=parse_working_days(["sun", "mon", "tue", "wed", "thu"]),
    )
    # Friday is excluded, Saturday too; Sunday 09:00 is first.
    assert slots[0].start == at(6, "09:00")


def test_per_day_limit_spreads_results():
    free = fb_from([])
    slots = find_slots(
        {"Ann": free}, {},
        window_start=at(0, "09:00"), window_end=at(3, "00:00"),
        duration=dt.timedelta(minutes=30), max_per_day=2, max_results=4,
    )
    assert [s.start for s in slots] == [at(0, "09:00"), at(0, "09:30"), at(1, "09:00"), at(1, "09:30")]


def test_tentative_and_optional_handling():
    ann = fb_from([("0+09:00", "10:00", "1")])
    cat = fb_from([("0+09:00", "09:30", "2")])
    strict = find_slots({"Ann": ann}, {}, window_start=at(0, "09:00"), window_end=at(0, "10:30"),
                        duration=dt.timedelta(minutes=30))
    assert strict[0].start == at(0, "10:00")
    loose = find_slots({"Ann": ann}, {"Cat": cat}, window_start=at(0, "09:00"), window_end=at(0, "10:30"),
                       duration=dt.timedelta(minutes=30), tentative_is_free=True)
    assert loose[0].start == at(0, "09:00")
    assert loose[0].tentative == ["Ann"] and loose[0].optional_conflicts == ["Cat"]


def test_unknown_data_blocks_slot():
    short = parse_free_busy("0" * 20, MON, 30)  # data ends at 10:00 on day 0
    slots = find_slots({"Ann": short}, {}, window_start=at(0, "09:00"), window_end=at(0, "12:00"),
                       duration=dt.timedelta(minutes=30))
    assert [s.start for s in slots] == [at(0, "09:00"), at(0, "09:30")]


def test_unaligned_start_rounds_up_and_parsers():
    slots = find_slots({"Ann": fb_from([])}, {}, window_start=at(0, "09:07"), window_end=at(0, "10:30"),
                       duration=dt.timedelta(minutes=30), step_minutes=15)
    assert slots[0].start == at(0, "09:15")
    assert parse_hhmm("08:30") == dt.time(8, 30)
    with pytest.raises(ValueError):
        parse_hhmm("8am")
    with pytest.raises(ValueError):
        parse_working_days(["funday"])
    assert parse_working_days(None) == {0, 1, 2, 3, 4}
    assert parse_working_days(["Sunday"]) == {6}


# ---------------------------------------------------------------- client


class _Recip:
    def __init__(self, name, fb=None, resolves=True, fb_error=False):
        self.Name = name
        self._fb = fb
        self._resolves = resolves
        self._fb_error = fb_error
        self.asked = []

    def Resolve(self):
        return self._resolves

    def FreeBusy(self, start, interval, complete):
        self.asked.append((start, interval, complete))
        if self._fb_error:
            raise Exception("no free/busy")
        return self._fb


class _NS:
    def __init__(self, recips, me=None, shared=None):
        self.recips = recips
        self.CurrentUser = me
        self.shared = shared

    def CreateRecipient(self, who):
        return self.recips[who]

    def GetSharedDefaultFolder(self, recip, kind):
        assert kind == 9
        if self.shared is None:
            raise Exception("access denied")
        return self.shared


def test_get_free_busy_reports_blocks_and_problems():
    raw = "0" * 18 + "22" + "0" * 28  # busy 09:00-10:00 at 30-min resolution
    ns = _NS({
        "ann@x.com": _Recip("Ann", raw),
        "ext@other.com": _Recip("ext@other.com", fb_error=True),
        "nobody": _Recip("nobody", resolves=False),
    })
    out = cal.get_free_busy(None, ns, people=["ann@x.com", "ext@other.com", "nobody"],
                            start="2026-11-09T08:00:00", end="2026-11-09T18:00:00")
    [ann] = out["people"]
    assert ann["name"] == "Ann" and len(ann["busy"]) == 1
    assert ann["busy"][0]["status"] == "busy"
    assert {p["who"] for p in out["problems"]} == {"ext@other.com", "nobody"}
    # Asked from local midnight, complete format, with to_com relabel.
    start, interval, complete = ns.recips["ann@x.com"].asked[0]
    assert start == dt.datetime(2026, 11, 9, tzinfo=dt.timezone.utc) and interval == 30 and complete


def test_find_meeting_times_end_to_end(monkeypatch):
    class _FixedDT(dt.datetime):
        @classmethod
        def now(cls, tz=None):
            return dt.datetime(2026, 11, 8, 12, 0)

    monkeypatch.setattr(cal.dt, "datetime", _FixedDT)
    day = ["0"] * 96  # 15-min resolution, one day
    busy_morning = day[:]
    for i in range(36, 44):  # 09:00-11:00
        busy_morning[i] = "2"
    raw = "".join(busy_morning) + "0" * 96 * 7
    ns = _NS({"ann@x.com": _Recip("Ann", raw)}, me=_Recip("Me", "0" * 96 * 8))
    out = cal.find_meeting_times(
        None, ns, attendees=["ann@x.com"], duration_minutes=60,
        start="2026-11-09T00:00:00", end="2026-11-10T00:00:00", max_per_day=2,
    )
    assert out["required"] == ["Me", "Ann"]
    assert [s["start"][:16] for s in out["slots"]] == ["2026-11-09T11:00", "2026-11-09T12:00"]


def test_find_meeting_times_validates_input():
    ns = _NS({})
    with pytest.raises(OutlookError, match="day_end"):
        cal.find_meeting_times(None, ns, attendees=["a"], day_start="17:00", day_end="09:00",
                               start="2030-01-01T00:00", end="2030-01-02T00:00")
    with pytest.raises(OutlookError, match="62 days"):
        cal.find_meeting_times(None, ns, attendees=["a"], start="2030-01-01T00:00", end="2030-06-01T00:00")


def test_other_calendar_needs_permission():
    ns = _NS({"ann@x.com": _Recip("Ann")}, shared=None)
    with pytest.raises(OutlookError, match="free/busy access"):
        cal.calendar_folder(ns, "ann@x.com")
    folder = object()
    ns2 = _NS({"ann@x.com": _Recip("Ann")}, shared=folder)
    assert cal.calendar_folder(ns2, "ann@x.com") == (folder, "Ann")
    with pytest.raises(OutlookError, match="address book"):
        cal.calendar_folder(_NS({"x": _Recip("x", resolves=False)}), "x")
