"""Free/busy parsing and meeting-slot search (the "scheduling assistant").

Outlook's ``Recipient.FreeBusy(start, interval, True)`` returns one digit
per ``interval`` minutes, starting at local midnight of ``start``:

    0 free · 1 tentative · 2 busy · 3 out of office · 4 working elsewhere

It reads the published free/busy data from Exchange, so it works for
anyone in the organization without permission to open their calendar.

Everything here is pure (no COM) so it can be unit-tested.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field

FREE, TENTATIVE, BUSY, OOF, ELSEWHERE = 0, 1, 2, 3, 4
UNKNOWN = -1

STATUS_NAMES = {
    FREE: "free",
    TENTATIVE: "tentative",
    BUSY: "busy",
    OOF: "out_of_office",
    ELSEWHERE: "working_elsewhere",
    UNKNOWN: "unknown",
}

DAY_NAMES = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]  # datetime.weekday() order


@dataclass
class FreeBusy:
    """Parsed free/busy for one person."""

    origin: dt.datetime  # local midnight the string starts at
    interval: int  # minutes per character
    codes: list[int]

    def status_at(self, when: dt.datetime) -> int:
        idx = int((when - self.origin).total_seconds() // (self.interval * 60))
        if 0 <= idx < len(self.codes):
            return self.codes[idx]
        return UNKNOWN

    def statuses_between(self, start: dt.datetime, end: dt.datetime) -> list[int]:
        """Statuses of every interval overlapping [start, end)."""
        out = []
        step = dt.timedelta(minutes=self.interval)
        first = int((start - self.origin).total_seconds() // (self.interval * 60))
        t = self.origin + first * step
        while t < end:
            out.append(self.status_at(t))
            t += step
        return out


def parse_free_busy(raw: str, origin: dt.datetime, interval: int) -> FreeBusy:
    codes = []
    for ch in raw or "":
        codes.append(int(ch) if ch.isdigit() else UNKNOWN)
    return FreeBusy(origin=origin, interval=interval, codes=codes)


def busy_blocks(
    fb: FreeBusy, start: dt.datetime, end: dt.datetime
) -> list[dict[str, object]]:
    """Merge consecutive non-free intervals within [start, end) into blocks."""
    blocks: list[dict[str, object]] = []
    step = dt.timedelta(minutes=fb.interval)
    first = int((start - fb.origin).total_seconds() // (fb.interval * 60))
    t = fb.origin + first * step
    current: dict[str, object] | None = None
    while t < end:
        code = fb.status_at(t)
        if code == FREE:
            current = None
        else:
            name = STATUS_NAMES.get(code, "unknown")
            if current is not None and current["status"] == name:
                current["end"] = min(t + step, end)
            else:
                current = {"start": max(t, start), "end": min(t + step, end), "status": name}
                blocks.append(current)
        t += step
    return blocks


def parse_working_days(days: list[str] | None) -> set[int]:
    if not days:
        return {0, 1, 2, 3, 4}
    out = set()
    for d in days:
        key = d.strip().lower()[:3]
        if key not in DAY_NAMES:
            raise ValueError(f"Unknown day '{d}'. Use mon, tue, wed, thu, fri, sat, sun.")
        out.add(DAY_NAMES.index(key))
    return out


def parse_hhmm(value: str) -> dt.time:
    try:
        h, m = value.strip().split(":")
        return dt.time(int(h), int(m))
    except Exception as exc:
        raise ValueError(f"Invalid time '{value}'. Use HH:MM, e.g. '09:00'.") from exc


@dataclass
class Slot:
    start: dt.datetime
    end: dt.datetime
    optional_conflicts: list[str] = field(default_factory=list)
    tentative: list[str] = field(default_factory=list)


def _blocked(code: int, tentative_is_free: bool) -> bool:
    if code in (FREE, ELSEWHERE):
        return False
    if code == TENTATIVE:
        return not tentative_is_free
    return True  # busy, OOF, unknown


def find_slots(
    required: dict[str, FreeBusy],
    optional: dict[str, FreeBusy],
    *,
    window_start: dt.datetime,
    window_end: dt.datetime,
    duration: dt.timedelta,
    step_minutes: int = 30,
    day_start: dt.time = dt.time(9, 0),
    day_end: dt.time = dt.time(17, 0),
    working_days: set[int] | None = None,
    tentative_is_free: bool = False,
    max_results: int = 10,
    max_per_day: int = 3,
) -> list[Slot]:
    """Find meeting slots where every required person is available.

    Candidate starts are aligned to ``step_minutes`` past the hour, lie
    inside working hours on working days, and the whole meeting must fit
    before ``day_end``. Slots are returned in time order, at most
    ``max_per_day`` per day so results spread across the window; optional
    attendees who are busy are listed per slot rather than excluding it.
    """
    days = working_days if working_days is not None else {0, 1, 2, 3, 4}
    step = dt.timedelta(minutes=step_minutes)
    results: list[Slot] = []

    # First aligned candidate at or after window_start.
    midnight = window_start.replace(hour=0, minute=0, second=0, microsecond=0)
    elapsed = window_start - midnight
    n = -(-int(elapsed.total_seconds()) // int(step.total_seconds()))  # ceil
    t = midnight + n * step

    per_day: dict[dt.date, int] = {}
    while t + duration <= window_end and len(results) < max_results:
        day_open = dt.datetime.combine(t.date(), day_start)
        day_close = dt.datetime.combine(t.date(), day_end)
        if t.weekday() not in days or t + duration > day_close:
            # Jump to the next day's opening time.
            nxt = dt.datetime.combine(t.date() + dt.timedelta(days=1), day_start)
            t = max(nxt, t + step)
            continue
        if t < day_open:
            t = day_open
            continue
        if per_day.get(t.date(), 0) >= max_per_day:
            t = dt.datetime.combine(t.date() + dt.timedelta(days=1), day_start)
            continue

        end = t + duration
        ok = True
        tentative: list[str] = []
        for name, fb in required.items():
            codes = fb.statuses_between(t, end)
            if any(_blocked(c, tentative_is_free) for c in codes):
                ok = False
                break
            if TENTATIVE in codes:
                tentative.append(name)
        if ok:
            conflicts = [
                name
                for name, fb in optional.items()
                if any(_blocked(c, tentative_is_free) for c in fb.statuses_between(t, end))
            ]
            results.append(Slot(t, end, conflicts, tentative))
            per_day[t.date()] = per_day.get(t.date(), 0) + 1
            # Don't offer overlapping alternatives: continue after this slot,
            # realigned to the step grid.
            t = end
            rem = (t - midnight) % step
            if rem:
                t += step - rem
            continue
        t += step
    return results
