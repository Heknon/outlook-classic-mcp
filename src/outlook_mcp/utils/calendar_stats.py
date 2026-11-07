"""Aggregate statistics over a calendar's events (pure, testable).

Input: event dicts with naive local ``start``/``end`` datetimes plus
optional ``subject``, ``organizer``, ``attendees`` (names), ``recurring``,
``private``, ``all_day``, ``categories`` (list), ``busy`` (status name).
Output: counts and distributions an agent can interpret — the arithmetic
is done here so it doesn't have to be done by the model.
"""

from __future__ import annotations

import datetime as dt
import re
from collections import Counter, defaultdict
from statistics import median
from typing import Any

WEEKDAYS = ["mon", "tue", "wed", "thu", "fri", "sat", "sun"]
_PREFIX = re.compile(r"^(?:(?:re|fw|fwd|aw|wg|updated|canceled|cancelled|accepted)\s*:\s*)+", re.I)


def _hours(td: dt.timedelta) -> float:
    return round(td.total_seconds() / 3600, 2)


def _norm_subject(s: str) -> str:
    s = _PREFIX.sub("", s or "").strip()
    return re.sub(r"\s+", " ", s)


def _hhmm(minutes: float) -> str:
    m = int(round(minutes))
    return f"{m // 60:02d}:{m % 60:02d}"


def analyze_events(
    events: list[dict[str, Any]],
    *,
    owner: str | None = None,
    start: dt.datetime,
    end: dt.datetime,
    top: int = 10,
) -> dict[str, Any]:
    timed = [e for e in events if not e.get("all_day") and e.get("busy", "busy") != "free"]
    all_day = [e for e in events if e.get("all_day")]
    days_in_range = max(1, (end.date() - start.date()).days)
    weeks = max(1.0, days_in_range / 7)

    total = dt.timedelta()
    by_weekday_count: Counter[str] = Counter()
    by_weekday_hours: dict[str, dt.timedelta] = defaultdict(dt.timedelta)
    by_hour: Counter[int] = Counter()
    by_week: dict[str, dt.timedelta] = defaultdict(dt.timedelta)
    by_week_count: Counter[str] = Counter()
    durations: list[float] = []
    per_day: dict[dt.date, list[dict[str, Any]]] = defaultdict(list)
    organizers: Counter[str] = Counter()
    attendees: Counter[str] = Counter()
    one_on_ones: Counter[str] = Counter()
    categories: Counter[str] = Counter()
    recurring_titles: dict[str, dict[str, Any]] = {}
    subjects: Counter[str] = Counter()
    private = 0
    owner_l = (owner or "").strip().lower()

    for e in timed:
        s, f = e["start"], e["end"]
        dur = max(f - s, dt.timedelta())
        total += dur
        durations.append(dur.total_seconds() / 60)
        wd = WEEKDAYS[s.weekday()]
        by_weekday_count[wd] += 1
        by_weekday_hours[wd] += dur
        by_hour[s.hour] += 1
        iso = s.isocalendar()
        week = f"{iso[0]}-W{iso[1]:02d}"
        by_week[week] += dur
        by_week_count[week] += 1
        per_day[s.date()].append(e)
        if e.get("private"):
            private += 1
            continue  # count the time, never the details
        org = (e.get("organizer") or "").strip()
        if org:
            organizers[org] += 1
        people = [a for a in (e.get("attendees") or []) if a and a.strip().lower() != owner_l]
        for a in people:
            attendees[a] += 1
        if len(people) == 1:
            one_on_ones[people[0]] += 1
        for c in e.get("categories") or []:
            categories[c] += 1
        title = _norm_subject(e.get("subject", ""))
        if title:
            subjects[title] += 1
            if e.get("recurring"):
                r = recurring_titles.setdefault(title, {"title": title, "count": 0, "weekdays": Counter(),
                                                        "times": Counter(), "minutes": []})
                r["count"] += 1
                r["weekdays"][wd] += 1
                r["times"][s.strftime("%H:%M")] += 1
                r["minutes"].append(dur.total_seconds() / 60)

    firsts = [min(d, key=lambda e: e["start"])["start"] for d in per_day.values()]
    lasts = [max(d, key=lambda e: e["end"])["end"] for d in per_day.values()]
    to_min = lambda t: t.hour * 60 + t.minute  # noqa: E731

    recurring = []
    for r in sorted(recurring_titles.values(), key=lambda r: -r["count"])[:top]:
        recurring.append({
            "title": r["title"],
            "occurrences": r["count"],
            "usual_days": [d for d, _ in r["weekdays"].most_common() if d],
            "usual_time": r["times"].most_common(1)[0][0],
            "minutes": round(median(r["minutes"])),
        })

    busiest = sorted(per_day.items(), key=lambda kv: -sum((e["end"] - e["start"]).total_seconds() for e in kv[1]))
    out: dict[str, Any] = {
        "owner": owner,
        "range": {"start": start.isoformat(), "end": end.isoformat(), "days": days_in_range},
        "meetings": len(timed),
        "meeting_hours": _hours(total),
        "hours_per_week": round(_hours(total) / weeks, 1),
        "meetings_per_week": round(len(timed) / weeks, 1),
        "days_with_meetings": len(per_day),
        "median_meeting_minutes": round(median(durations)) if durations else 0,
        "by_weekday": [
            {"day": d, "meetings": by_weekday_count[d], "hours": _hours(by_weekday_hours[d])}
            for d in WEEKDAYS if by_weekday_count[d]
        ],
        "by_start_hour": [{"hour": f"{h:02d}:00", "meetings": n} for h, n in sorted(by_hour.items())],
        "typical_day": {
            "first_meeting_starts": _hhmm(median([to_min(t) for t in firsts])) if firsts else None,
            "last_meeting_ends": _hhmm(median([to_min(t) for t in lasts])) if lasts else None,
        },
        "by_week": [{"week": w, "meetings": by_week_count[w], "hours": _hours(h)} for w, h in sorted(by_week.items())],
        "busiest_days": [
            {"date": d.isoformat(), "meetings": len(evs),
             "hours": _hours(sum((e["end"] - e["start"] for e in evs), dt.timedelta()))}
            for d, evs in busiest[:5]
        ],
        "recurring_meetings": recurring,
        "most_frequent_titles": [{"title": t, "count": n} for t, n in subjects.most_common(top)],
        "top_people": [{"name": n, "meetings": c} for n, c in attendees.most_common(top)],
        "top_one_on_ones": [{"name": n, "meetings": c} for n, c in one_on_ones.most_common(top)],
        "top_organizers": [{"name": n, "meetings": c} for n, c in organizers.most_common(top)],
        "categories": [{"name": n, "meetings": c} for n, c in categories.most_common(top)],
        "all_day_events": len(all_day),
        "private_appointments": private,
    }
    if owner_l and organizers:
        mine = sum(c for n, c in organizers.items() if n.strip().lower() == owner_l)
        out["organized_by_owner"] = mine
    return out
