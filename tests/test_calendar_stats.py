"""Calendar analytics: exact counts on a known week pattern."""

import datetime as dt

from outlook_mcp.utils.calendar_stats import analyze_events

SUN = dt.datetime(2026, 9, 6)  # a Sunday


def ev(day, h, m, minutes, subject, attendees=(), recurring=True, private=False, organizer="Dana Levi", **kw):
    s = SUN + dt.timedelta(days=day, hours=h, minutes=m)
    return {"start": s, "end": s + dt.timedelta(minutes=minutes), "subject": subject,
            "attendees": list(attendees), "recurring": recurring, "private": private,
            "organizer": organizer, **kw}


def two_weeks():
    out = []
    for week in (0, 7):
        for d in range(5):  # Sun..Thu
            out.append(ev(week + d, 9, 30, 15, "Product standup", ["Dana Levi", "Yossi", "Cohen"]))
        out.append(ev(week + 1, 11, 0, 30, "1:1 Dana / Yossi", ["Yossi"]))
        out.append(ev(week + 3, 14, 0, 60, "Roadmap review", ["Ori", "Yossi", "Cohen"], organizer="Ori"))
        out.append(ev(week + 4, 7, 30, 60, "Pilates", private=True, subject_hidden=True))
    out.append(ev(2, 0, 0, 1440, "Conference", all_day=True))
    out.append(ev(2, 12, 0, 60, "Lunch (free)", busy="free"))
    return out


def test_counts_and_load():
    out = analyze_events(two_weeks(), owner="Dana Levi", start=SUN, end=SUN + dt.timedelta(days=14))
    assert out["meetings"] == 16  # 10 standups + 2 1:1 + 2 roadmap + 2 private; all-day and free excluded
    assert out["meeting_hours"] == round((10 * 15 + 2 * 30 + 2 * 60 + 2 * 60) / 60, 2)
    assert out["meetings_per_week"] == 8.0
    assert out["all_day_events"] == 1 and out["private_appointments"] == 2
    days = {d["day"]: d["meetings"] for d in out["by_weekday"]}
    assert days == {"sun": 2, "mon": 4, "tue": 2, "wed": 4, "thu": 4}
    assert out["typical_day"]["first_meeting_starts"] == "09:30"


def test_routine_people_and_privacy():
    out = analyze_events(two_weeks(), owner="Dana Levi", start=SUN, end=SUN + dt.timedelta(days=14))
    rec = {r["title"]: r for r in out["recurring_meetings"]}
    assert rec["Product standup"]["occurrences"] == 10 and rec["Product standup"]["usual_time"] == "09:30"
    assert rec["1:1 Dana / Yossi"]["usual_days"] == ["mon"] and rec["Roadmap review"]["minutes"] == 60
    assert "Pilates" not in str(out)  # private: counted, never named
    assert out["top_people"][0] == {"name": "Yossi", "meetings": 14}
    assert out["top_one_on_ones"] == [{"name": "Yossi", "meetings": 2}]
    assert out["organized_by_owner"] == 12
    assert {o["name"] for o in out["top_organizers"]} == {"Dana Levi", "Ori"}


def test_empty_range():
    out = analyze_events([], owner="x", start=SUN, end=SUN + dt.timedelta(days=7))
    assert out["meetings"] == 0 and out["typical_day"]["first_meeting_starts"] is None
