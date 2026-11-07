"""Calendar (appointment / meeting) COM operations."""

from __future__ import annotations

import datetime as dt
from typing import Any

from outlook_mcp.client.folders import _safe_get, get_item_by_id
from outlook_mcp.constants import (
    OL_APPOINTMENT_ITEM,
    OL_FOLDER_CALENDAR,
    OL_MEETING,
    OL_MEETING_ACCEPTED,
    OL_MEETING_DECLINED,
    OL_MEETING_TENTATIVE,
    OL_RECURS_DAILY,
    OL_RECURS_MONTHLY,
    OL_RECURS_WEEKLY,
    OL_APPT_MASTER,
    OL_MEETING_CANCELED,
    OL_OPTIONAL,
    OL_RECURS_YEARLY,
    OL_REQUIRED,
    OL_RESPONSE_ORGANIZED,
)
from outlook_mcp.errors import OutlookError
from outlook_mcp.schemas import Recurrence
from outlook_mcp.utils.formatting import from_iso, to_com, to_dasl_date, to_iso, truncate
from outlook_mcp.utils.policy import DRAFT_ONLY_NOTE, draft_only
from outlook_mcp.utils.calendar_stats import analyze_events
from outlook_mcp.utils.scheduling import (
    FreeBusy,
    busy_blocks,
    find_slots,
    parse_free_busy,
    parse_hhmm,
    parse_working_days,
)

# How far before `start` to widen the Restrict() lower bound so events that
# begin earlier but still overlap the requested window (multi-day / all-day
# appointments) aren't excluded. Outlook's DASL Restrict() is only reliable
# ANDing two bounds on the *same* property (verified live), so we can't add
# a "[End] > start" clause directly — instead we over-fetch on [Start] and
# apply the precise overlap check in Python below.
_OVERLAP_LOOKBACK = dt.timedelta(days=3)


_RECURRENCE_TYPE_MAP = {
    "daily": OL_RECURS_DAILY,
    "weekly": OL_RECURS_WEEKLY,
    "monthly": OL_RECURS_MONTHLY,
    "yearly": OL_RECURS_YEARLY,
}


_RECURRENCE_STATE_NAMES = {0: "none", 1: "series", 2: "occurrence", 3: "exception"}


def _event_summary(ev: Any) -> dict[str, Any]:
    recur_state = _safe_get(ev, "RecurrenceState", 0)
    return {
        "entry_id": _safe_get(ev, "EntryID"),
        "subject": _safe_get(ev, "Subject", ""),
        "start": to_iso(_safe_get(ev, "Start")),
        "end": to_iso(_safe_get(ev, "End")),
        "location": _safe_get(ev, "Location", ""),
        "organizer": _safe_get(ev, "Organizer", ""),
        "is_recurring": bool(recur_state),
        # For an occurrence of a recurring series, entry_id is the SERIES id;
        # pass this event's "start" as occurrence_start to act on just it.
        "recurrence": _RECURRENCE_STATE_NAMES.get(recur_state, "none"),
        "all_day": bool(_safe_get(ev, "AllDayEvent", False)),
        "preview": truncate(_safe_get(ev, "Body", ""), 200),
    }


def _event_full(ev: Any) -> dict[str, Any]:
    attendees = []
    if _safe_get(ev, "Recipients"):
        for r in ev.Recipients:
            attendees.append(
                {
                    "name": _safe_get(r, "Name"),
                    "address": _safe_get(r, "Address"),
                    "type": _safe_get(r, "Type"),
                    "response": _safe_get(r, "MeetingResponseStatus"),
                }
            )
    return {
        **_event_summary(ev),
        "body": _safe_get(ev, "Body", ""),
        "attendees": attendees,
        "reminder_minutes": _safe_get(ev, "ReminderMinutesBeforeStart"),
        "categories": _safe_get(ev, "Categories", ""),
    }


def _apply_recurrence(appt: Any, spec: Recurrence) -> None:
    pattern = appt.GetRecurrencePattern()
    pattern.RecurrenceType = _RECURRENCE_TYPE_MAP[spec.type]
    pattern.Interval = spec.interval
    if spec.occurrences is not None:
        pattern.Occurrences = spec.occurrences
    if spec.end_date is not None:
        pattern.PatternEndDate = to_com(from_iso(spec.end_date))


def _as_naive(value: Any) -> dt.datetime | None:
    if not hasattr(value, "year"):
        return None
    return dt.datetime(
        value.year, value.month, value.day, value.hour, value.minute, value.second
    )


def _events_between(
    cal: Any, start_dt: dt.datetime, end_dt: dt.datetime, limit: int, include_recurrences: bool = True
) -> tuple[list[Any], bool]:
    """Items of calendar folder ``cal`` overlapping [start_dt, end_dt)."""
    items = cal.Items
    # Outlook requires the collection sorted ascending by [Start] BEFORE
    # IncludeRecurrences is set, or recurring series expand incorrectly.
    items.Sort("[Start]")
    if include_recurrences:
        items.IncludeRecurrences = True

    # DASL property query (locale-independent ISO dates) instead of a plain
    # Jet literal — see to_dasl_date() for why the plain "[Start] >= 'mm/dd
    # /yyyy ...'" form silently returns the wrong events on non-US locales.
    restrict = (
        f"@SQL=\"urn:schemas:calendar:dtstart\" >= "
        f"'{to_dasl_date(start_dt - _OVERLAP_LOOKBACK)}' AND "
        f"\"urn:schemas:calendar:dtstart\" <= '{to_dasl_date(end_dt)}'"
    )
    filtered = items.Restrict(restrict)

    found: list[Any] = []
    truncated = False
    for ev in filtered:
        ev_start = _as_naive(_safe_get(ev, "Start"))
        ev_end = _as_naive(_safe_get(ev, "End"))
        if ev_start is None:
            continue
        # Restrict() above only bounds [Start]; apply the real overlap test
        # (and, for un-expanded recurring masters, the real range check —
        # Restrict() does not reliably filter those by date, verified live:
        # masters years outside the range still pass through when
        # IncludeRecurrences is left off).
        if ev_start > end_dt:
            if include_recurrences:
                # Sorted ascending by [Start]: nothing later can overlap, and
                # an expanded recurrence collection is unbounded — stop here.
                break
            continue
        if ev_end is not None and ev_end <= start_dt:
            continue
        if ev_end is None and ev_start < start_dt:
            continue
        if len(found) >= limit:
            truncated = True
            break
        found.append(ev)
    return found, truncated


def list_events(
    outlook: Any,
    namespace: Any,
    *,
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    include_recurrences: bool = True,
    mailbox: str | None = None,
    offset: int = 0,
    compact: bool = False,
    include_attendees: bool = False,
) -> dict[str, Any]:
    """Events overlapping [start, end) — past ranges work too (history)."""
    cal, owner = calendar_folder(namespace, mailbox)
    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))
    if end_dt <= start_dt:
        raise OutlookError("end must be after start.")
    found, truncated = _events_between(cal, start_dt, end_dt, offset + limit, include_recurrences)
    page = found[offset:]
    results = []
    for ev in page:
        row = _compact_event(ev) if compact else _event_summary(ev)
        if include_attendees:
            row["attendees"] = _attendee_names(ev)
        if owner:
            row = _mask_private(row, ev)
        results.append(row)
    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "offset": offset,
        "items": results,
        "truncated": truncated,
        "next_offset": offset + len(results) if truncated else None,
        **({"mailbox": owner} if owner else {}),
    }


def _attendee_names(ev: Any) -> list[str]:
    names = []
    recips = _safe_get(ev, "Recipients")
    for r in list(recips or []):
        n = _safe_get(r, "Name")
        if n:
            names.append(str(n))
    return names


def _compact_event(ev: Any) -> dict[str, Any]:
    recur_state = _safe_get(ev, "RecurrenceState", 0)
    row = {
        "entry_id": _safe_get(ev, "EntryID"),
        "start": to_iso(_safe_get(ev, "Start")),
        "end": to_iso(_safe_get(ev, "End")),
        "subject": _safe_get(ev, "Subject", ""),
        "organizer": _safe_get(ev, "Organizer", ""),
        "recurrence": _RECURRENCE_STATE_NAMES.get(recur_state, "none"),
    }
    if _safe_get(ev, "AllDayEvent", False):
        row["all_day"] = True
    cats = _safe_get(ev, "Categories", "")
    if cats:
        row["categories"] = cats
    return row


# OlSensitivity
_PRIVATE = 2
_BUSY_NAMES = {0: "free", 1: "tentative", 2: "busy", 3: "out_of_office", 4: "working_elsewhere"}


def _mask_private(row: dict[str, Any], ev: Any) -> dict[str, Any]:
    """Someone else's private appointment: keep the time, hide the details."""
    if _safe_get(ev, "Sensitivity", 0) == _PRIVATE:
        return {
            **{k: row.get(k) for k in ("entry_id", "start", "end", "all_day", "is_recurring", "recurrence")},
            "subject": "Private appointment",
            "private": True,
        }
    return row


def _stat_row(ev: Any, mask_private: bool) -> dict[str, Any] | None:
    s, e = _as_naive(_safe_get(ev, "Start")), _as_naive(_safe_get(ev, "End"))
    if s is None or e is None:
        return None
    private = mask_private and _safe_get(ev, "Sensitivity", 0) == _PRIVATE
    cats = str(_safe_get(ev, "Categories", "") or "")
    return {
        "start": s,
        "end": e,
        "subject": "" if private else str(_safe_get(ev, "Subject", "") or ""),
        "organizer": "" if private else str(_safe_get(ev, "Organizer", "") or ""),
        "attendees": [] if private else _attendee_names(ev),
        "recurring": _safe_get(ev, "RecurrenceState", 0) in (1, 2, 3) or bool(_safe_get(ev, "IsRecurring", False)),
        "private": private,
        "all_day": bool(_safe_get(ev, "AllDayEvent", False)),
        "categories": [c.strip() for c in cats.replace(";", ",").split(",") if c.strip()],
        "busy": _BUSY_NAMES.get(_safe_get(ev, "BusyStatus", 2), "busy"),
    }


def calendar_stats(
    outlook: Any,
    namespace: Any,
    *,
    mailbox: str | None = None,
    start: str | None = None,
    end: str | None = None,
    max_events: int = 3000,
) -> dict[str, Any]:
    """Patterns in someone's calendar over a period (default: last 8 weeks).

    Uses the calendar itself when it's yours or shared with you; otherwise
    falls back to free/busy (times only — no titles or people).
    """
    now = dt.datetime.now().replace(second=0, microsecond=0)
    end_dt = from_iso(end) or now
    start_dt = from_iso(start) or (end_dt - dt.timedelta(weeks=8))
    if end_dt <= start_dt:
        raise OutlookError("end must be after start.")
    if (end_dt - start_dt).days > 400:
        raise OutlookError("Range too long: use at most about a year.")
    if mailbox:
        recip = resolve_recipient(namespace, mailbox)
        owner = str(_safe_get(recip, "Name", mailbox))
        try:
            cal = namespace.GetSharedDefaultFolder(recip, OL_FOLDER_CALENDAR)
        except Exception:
            cal = None
    else:
        recip = _safe_get(namespace, "CurrentUser")
        owner = str(_safe_get(recip, "Name", "") or "me")
        cal = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)

    if cal is not None:
        found, truncated = _events_between(cal, start_dt, end_dt, max_events)
        rows = [r for r in (_stat_row(ev, mask_private=bool(mailbox)) for ev in found) if r]
        out = analyze_events(rows, owner=owner, start=start_dt, end=end_dt)
        out["access"] = "calendar"
        if truncated:
            out["truncated"] = True
            out["note"] = f"Stopped at {max_events} events; shorten the range for complete numbers."
        return out

    # Free/busy only: busy blocks become untitled "events".
    fb = _fetch_free_busy(recip, start_dt, 30)
    blocks = busy_blocks(fb, start_dt, end_dt)
    rows = [{"start": b["start"], "end": b["end"], "busy": b["status"]} for b in blocks
            if b["status"] in ("busy", "tentative", "out_of_office")]
    out = analyze_events(rows, owner=owner, start=start_dt, end=end_dt)
    out["access"] = "free_busy_only"
    out["note"] = (
        f"{owner}'s calendar isn't shared with you, so only busy times are known (no titles, "
        "people or recurring meetings). Exchange publishes limited past free/busy, so older "
        "weeks may show as empty."
    )
    return out


def _meetings_for(namespace: Any, recip: Any, start_dt: dt.datetime, end_dt: dt.datetime,
                  is_me: bool = False, limit: int = 50) -> list[dict[str, Any]] | None:
    """Meeting titles in the window, or None when the calendar can't be opened."""
    try:
        if is_me:
            cal = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
        else:
            cal = namespace.GetSharedDefaultFolder(recip, OL_FOLDER_CALENDAR)
        found, _ = _events_between(cal, start_dt, end_dt, limit)
    except Exception:
        return None
    out = []
    for ev in found:
        private = (not is_me) and _safe_get(ev, "Sensitivity", 0) == _PRIVATE
        busy = _BUSY_NAMES.get(_safe_get(ev, "BusyStatus", 2), "busy")
        if busy == "free" and not is_me:
            continue
        row = {
            "start": to_iso(_safe_get(ev, "Start")),
            "end": to_iso(_safe_get(ev, "End")),
            "subject": "Private appointment" if private else _safe_get(ev, "Subject", ""),
            "status": busy,
        }
        if not private:
            loc = _safe_get(ev, "Location", "")
            if loc:
                row["location"] = loc
        out.append(row)
    return out


def get_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    occurrence_start: str | None = None,
    mailbox: str | None = None,
) -> dict[str, Any]:
    store_id = None
    if mailbox:
        folder, _ = calendar_folder(namespace, mailbox)
        store_id = _safe_get(folder, "StoreID")
    return _event_full(_target_event(namespace, entry_id, occurrence_start, store_id))


def resolve_recipient(namespace: Any, who: str) -> Any:
    """Resolve a name or address against the address book (GAL)."""
    recip = namespace.CreateRecipient(who)
    try:
        ok = recip.Resolve()
    except Exception:
        ok = False
    if not ok:
        raise OutlookError(
            f"Couldn't find '{who}' in the address book. Use a full email address "
            "or check the name with outlook_resolve_name."
        )
    return recip


def calendar_folder(namespace: Any, mailbox: str | None) -> tuple[Any, str | None]:
    """Your calendar, or someone else's shared calendar.

    Opening another person's calendar needs at least Reviewer (read)
    permission on it, or a delegate grant. With only the default
    free/busy permission it fails — free/busy itself still works.
    """
    if not mailbox:
        return namespace.GetDefaultFolder(OL_FOLDER_CALENDAR), None
    recip = resolve_recipient(namespace, mailbox)
    try:
        folder = namespace.GetSharedDefaultFolder(recip, OL_FOLDER_CALENDAR)
    except Exception as exc:
        raise OutlookError(
            f"Can't open {_safe_get(recip, 'Name', mailbox)}'s calendar — you probably "
            "only have free/busy access. Use outlook_get_free_busy or "
            "outlook_find_meeting_times instead, or ask them to share their "
            "calendar with you (Reviewer permission)."
        ) from exc
    return folder, str(_safe_get(recip, "Name", mailbox))


def _midnight(value: dt.datetime) -> dt.datetime:
    return value.replace(hour=0, minute=0, second=0, microsecond=0)


def _fetch_free_busy(recip: Any, start: dt.datetime, interval: int) -> FreeBusy:
    origin = _midnight(start)
    raw = recip.FreeBusy(to_com(origin), interval, True)
    return parse_free_busy(str(raw or ""), origin, interval)


def _people_free_busy(
    namespace: Any,
    people: list[str],
    start: dt.datetime,
    interval: int,
    recips: dict[str, Any] | None = None,
) -> tuple[dict[str, FreeBusy], list[dict[str, str]]]:
    found: dict[str, FreeBusy] = {}
    problems: list[dict[str, str]] = []
    for who in people:
        try:
            recip = resolve_recipient(namespace, who)
        except OutlookError as exc:
            problems.append({"who": who, "error": str(exc)})
            continue
        name = str(_safe_get(recip, "Name", who) or who)
        if recips is not None:
            recips[name] = recip
        try:
            found[name] = _fetch_free_busy(recip, start, interval)
        except Exception:
            problems.append(
                {
                    "who": who,
                    "error": "No free/busy information available (external address, "
                    "or not published).",
                }
            )
    return found, problems


def _me_free_busy(namespace: Any, start: dt.datetime, interval: int) -> tuple[str, FreeBusy] | None:
    me = _safe_get(namespace, "CurrentUser")
    if me is None:
        return None
    try:
        return str(_safe_get(me, "Name", "me") or "me"), _fetch_free_busy(me, start, interval)
    except Exception:
        return None


def _range(start: str | None, end: str | None, default_days: int) -> tuple[dt.datetime, dt.datetime]:
    start_dt = from_iso(start) or dt.datetime.now().replace(second=0, microsecond=0)
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=default_days))
    if end_dt <= start_dt:
        raise OutlookError("end must be after start.")
    if (end_dt - start_dt).days > 62:
        raise OutlookError("Range too long: free/busy covers about two months; use at most 62 days.")
    return start_dt, end_dt


def _calendar_details(namespace: Any, recips: dict[str, Any], me_name: str | None,
                      names: list[str], start_dt: dt.datetime, end_dt: dt.datetime) -> list[dict[str, Any]]:
    """Per person: their meeting titles in the window if their calendar is
    visible to the user, else just the access level."""
    out = []
    for name in names:
        meetings = _meetings_for(namespace, recips.get(name), start_dt, end_dt, is_me=name == me_name, limit=30)
        if meetings is None:
            out.append({"name": name, "access": "free_busy_only"})
        else:
            out.append({"name": name, "access": "calendar", "meetings": meetings})
    return out


def get_free_busy(
    outlook: Any,
    namespace: Any,
    *,
    people: list[str],
    start: str | None = None,
    end: str | None = None,
    interval_minutes: int = 30,
    include_me: bool = False,
    include_titles: bool = True,
) -> dict[str, Any]:
    start_dt, end_dt = _range(start, end, 1)
    recips: dict[str, Any] = {}
    found, problems = _people_free_busy(namespace, people, start_dt, interval_minutes, recips)
    me_name = None
    if include_me:
        mine = _me_free_busy(namespace, start_dt, interval_minutes)
        if mine:
            me_name = mine[0]
            found = {mine[0]: mine[1], **found}
    result = []
    for name, fb in found.items():
        blocks = busy_blocks(fb, start_dt, end_dt)
        row: dict[str, Any] = {
            "name": name,
            "busy": [
                {"start": to_iso(b["start"]), "end": to_iso(b["end"]), "status": b["status"]}
                for b in blocks
            ],
        }
        if include_titles:
            meetings = _meetings_for(namespace, recips.get(name), start_dt, end_dt, is_me=name == me_name)
            if meetings is None:
                row["access"] = "free_busy_only"
            else:
                row["access"] = "calendar"
                row["meetings"] = meetings
        result.append(row)
    out: dict[str, Any] = {
        "start": to_iso(start_dt),
        "end": to_iso(end_dt),
        "interval_minutes": interval_minutes,
        "people": result,
    }
    if include_titles and any(r.get("access") == "free_busy_only" for r in result):
        out["note"] = (
            "Meeting titles are shown only for calendars shared with you; others show busy/free."
        )
    if problems:
        out["problems"] = problems
    return out


def find_meeting_times(
    outlook: Any,
    namespace: Any,
    *,
    attendees: list[str],
    duration_minutes: int = 30,
    optional_attendees: list[str] | None = None,
    start: str | None = None,
    end: str | None = None,
    day_start: str = "09:00",
    day_end: str = "17:00",
    working_days: list[str] | None = None,
    step_minutes: int = 30,
    include_me: bool = True,
    tentative_is_free: bool = False,
    max_results: int = 10,
    max_per_day: int = 3,
    include_titles: bool = True,
) -> dict[str, Any]:
    start_dt, end_dt = _range(start, end, 7)
    now = dt.datetime.now()
    if start_dt < now:
        start_dt = now
    try:
        days = parse_working_days(working_days)
        open_t, close_t = parse_hhmm(day_start), parse_hhmm(day_end)
    except ValueError as exc:
        raise OutlookError(str(exc)) from exc
    if close_t <= open_t:
        raise OutlookError("day_end must be after day_start.")

    # Free/busy resolution: fine enough for the step, and dividing an hour.
    interval = 15 if step_minutes % 15 == 0 else 5
    recips: dict[str, Any] = {}
    required, problems = _people_free_busy(namespace, attendees, start_dt, interval, recips)
    me_name = None
    if include_me:
        mine = _me_free_busy(namespace, start_dt, interval)
        if mine:
            me_name = mine[0]
            required = {mine[0]: mine[1], **required}
    optional, opt_problems = _people_free_busy(namespace, optional_attendees or [], start_dt, interval, recips)
    problems += opt_problems

    slots = find_slots(
        required,
        optional,
        window_start=start_dt,
        window_end=end_dt,
        duration=dt.timedelta(minutes=duration_minutes),
        step_minutes=step_minutes,
        day_start=open_t,
        day_end=close_t,
        working_days=days,
        tentative_is_free=tentative_is_free,
        max_results=max_results,
        max_per_day=max_per_day,
    )
    out: dict[str, Any] = {
        "duration_minutes": duration_minutes,
        "searched": {"start": to_iso(start_dt), "end": to_iso(end_dt)},
        "required": list(required),
        "optional": list(optional),
        **({"calendars": _calendar_details(namespace, recips, me_name, list(required) + list(optional),
                                           start_dt, end_dt)} if include_titles else {}),
        "count": len(slots),
        "slots": [
            {
                "start": to_iso(sl.start),
                "end": to_iso(sl.end),
                **({"optional_busy": sl.optional_conflicts} if sl.optional_conflicts else {}),
                **({"tentative": sl.tentative} if sl.tentative else {}),
            }
            for sl in slots
        ],
    }
    notes: list[str] = []
    if problems:
        out["problems"] = problems
        notes.append(
            "People listed under problems were left out of the search; check "
            "their availability another way."
        )
    if not slots:
        notes.append(
            "No time works for everyone in that window. Try a longer range, a "
            "shorter meeting, tentative_is_free=true, or moving someone to optional."
        )
    if notes:
        out["notes"] = notes
    return out


def _target_event(
    namespace: Any, entry_id: str, occurrence_start: str | None, store_id: str | None = None
) -> Any:
    """Resolve an event, or one occurrence of a recurring series.

    ``list_events`` returns the *series* EntryID for every occurrence of a
    recurring meeting, so acting on that id alone would hit the whole
    series. ``occurrence_start`` (the occurrence's original start, as
    listed) selects the single occurrence instead.
    """
    ev = get_item_by_id(namespace, entry_id, store_id)
    if not occurrence_start:
        return ev
    if not _safe_get(ev, "IsRecurring", False):
        return ev
    try:
        return ev.GetRecurrencePattern().GetOccurrence(to_com(from_iso(occurrence_start)))
    except Exception as exc:
        raise OutlookError(
            f"No occurrence of this series starts at {occurrence_start}. Pass the "
            "occurrence's 'start' exactly as outlook_list_events returned it."
        ) from exc


def _guard_series(ev: Any, apply_to_series: bool, verb: str) -> None:
    if _safe_get(ev, "RecurrenceState", 0) == OL_APPT_MASTER and not apply_to_series:
        raise OutlookError(
            f"This is a recurring series. To {verb} one occurrence pass "
            "occurrence_start (the occurrence's start from outlook_list_events); "
            f"to {verb} the whole series pass apply_to_series=true."
        )


def _is_meeting(ev: Any) -> bool:
    return _safe_get(ev, "MeetingStatus", 0) not in (0, None)


def _is_organizer(ev: Any) -> bool:
    return _safe_get(ev, "ResponseStatus") == OL_RESPONSE_ORGANIZED


def _all_day_bounds(start: str, end: str | None) -> tuple[dt.datetime, dt.datetime]:
    s = from_iso(start)
    s = dt.datetime(s.year, s.month, s.day)
    if end:
        e = from_iso(end)
        e = dt.datetime(e.year, e.month, e.day)
        # Accept an inclusive last day ("2026-05-01" to "2026-05-01").
        if e <= s:
            e = e + dt.timedelta(days=1)
    else:
        e = s + dt.timedelta(days=1)
    return s, e


def create_event(
    outlook: Any,
    namespace: Any,
    *,
    subject: str,
    start: str,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    attendees: list[str] | None = None,
    optional_attendees: list[str] | None = None,
    is_online_meeting: bool = False,
    reminder_minutes: int | None = 15,
    recurrence: Recurrence | None = None,
    all_day: bool = False,
    send_invites: bool = True,
    busy_status: str | None = None,
    categories: list[str] | None = None,
    category_color: str | None = None,
) -> dict[str, Any]:
    appt = outlook.CreateItem(OL_APPOINTMENT_ITEM)
    appt.Subject = subject
    if all_day:
        s, e = _all_day_bounds(start, end)
        appt.AllDayEvent = True
        appt.Start = to_com(s)
        appt.End = to_com(e)
    else:
        if not end:
            raise OutlookError("end is required unless all_day=true.")
        # to_com, not the bare datetime: pywin32 would read a naive value
        # as local and hand Outlook its UTC clock, shifting the event one
        # offset early.
        appt.Start = to_com(from_iso(start))
        appt.End = to_com(from_iso(end))
    if location:
        appt.Location = location
    if body:
        appt.Body = body
    if busy_status:
        appt.BusyStatus = _busy_code(busy_status)
    if categories:
        _apply_categories(namespace, appt, categories, category_color)
    if reminder_minutes is not None:
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = reminder_minutes
    else:
        appt.ReminderSet = False
    has_attendees = bool(attendees or optional_attendees)
    if has_attendees:
        appt.MeetingStatus = OL_MEETING
        for addr in attendees or []:
            appt.Recipients.Add(addr).Type = OL_REQUIRED
        for addr in optional_attendees or []:
            appt.Recipients.Add(addr).Type = OL_OPTIONAL
        appt.Recipients.ResolveAll()
    if recurrence is not None:
        _apply_recurrence(appt, recurrence)
    appt.Save()

    notes: list[str] = []
    sent = False
    if has_attendees:
        if send_invites and not draft_only():
            appt.Send()
            sent = True
        elif draft_only():
            notes.append(DRAFT_ONLY_NOTE)
        else:
            notes.append(
                "Saved in your calendar without sending the invitation. Open it "
                "in Outlook and click Send when ready."
            )
    out = {
        "status": "created",
        "entry_id": appt.EntryID,
        "subject": appt.Subject,
        "start": to_iso(appt.Start),
        "end": to_iso(appt.End),
        "all_day": bool(all_day),
        "invites_sent": sent,
    }
    if notes:
        out["notes"] = notes
    return out


_BUSY_MAP = {"free": 0, "tentative": 1, "busy": 2, "out_of_office": 3, "oof": 3, "working_elsewhere": 4}


def _busy_code(value: str) -> int:
    code = _BUSY_MAP.get(value.strip().lower())
    if code is None:
        raise OutlookError(
            "busy_status must be one of: free, tentative, busy, out_of_office, working_elsewhere."
        )
    return code


def _apply_categories(namespace: Any, item: Any, names: list[str], color: str | None) -> None:
    from outlook_mcp.client.categories import ensure_categories

    names = [n.strip() for n in names if n and n.strip()]
    ensure_categories(namespace, names, color)
    item.Categories = ", ".join(names)


def _remove_attendees(ev: Any, addresses: list[str]) -> list[str]:
    wanted = {a.strip().lower() for a in addresses if a}
    removed: list[str] = []
    recips = ev.Recipients
    for i in range(recips.Count, 0, -1):  # backwards: Remove shifts indexes
        r = recips.Item(i)
        keys = {
            str(_safe_get(r, "Address", "") or "").lower(),
            str(_safe_get(r, "Name", "") or "").lower(),
        }
        try:
            smtp = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress
            keys.add(str(smtp).lower())
        except Exception:
            pass
        if keys & wanted:
            removed.append(str(_safe_get(r, "Name", "")))
            recips.Remove(i)
    return removed


def update_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    subject: str | None = None,
    start: str | None = None,
    end: str | None = None,
    location: str | None = None,
    body: str | None = None,
    all_day: bool | None = None,
    reminder_minutes: int | None = None,
    busy_status: str | None = None,
    add_attendees: list[str] | None = None,
    add_optional_attendees: list[str] | None = None,
    remove_attendees: list[str] | None = None,
    send_update: bool = False,
    occurrence_start: str | None = None,
    apply_to_series: bool = False,
    categories: list[str] | None = None,
    category_color: str | None = None,
) -> dict[str, Any]:
    ev = _target_event(namespace, entry_id, occurrence_start)
    _guard_series(ev, apply_to_series, "change")
    if subject is not None:
        ev.Subject = subject
    if all_day is not None:
        ev.AllDayEvent = bool(all_day)
    if all_day and start is not None:
        s, e = _all_day_bounds(start, end)
        ev.Start = to_com(s)
        ev.End = to_com(e)
    else:
        if start is not None:
            ev.Start = to_com(from_iso(start))
        if end is not None:
            ev.End = to_com(from_iso(end))
    if location is not None:
        ev.Location = location
    if body is not None:
        ev.Body = body
    if reminder_minutes is not None:
        ev.ReminderSet = True
        ev.ReminderMinutesBeforeStart = reminder_minutes
    if busy_status:
        ev.BusyStatus = _busy_code(busy_status)
    if categories is not None:
        _apply_categories(namespace, ev, categories, category_color)

    removed: list[str] = []
    if add_attendees or add_optional_attendees or remove_attendees:
        if not _is_meeting(ev):
            ev.MeetingStatus = OL_MEETING
        for addr in add_attendees or []:
            ev.Recipients.Add(addr).Type = OL_REQUIRED
        for addr in add_optional_attendees or []:
            ev.Recipients.Add(addr).Type = OL_OPTIONAL
        if remove_attendees:
            removed = _remove_attendees(ev, remove_attendees)
        ev.Recipients.ResolveAll()
    ev.Save()

    notes: list[str] = []
    sent = False
    if send_update:
        if not _is_meeting(ev):
            notes.append("Not a meeting (no attendees), so there was no update to send.")
        elif not _is_organizer(ev):
            notes.append("You are not the organizer, so you can't send an update.")
        elif draft_only():
            notes.append(DRAFT_ONLY_NOTE)
        else:
            ev.Send()
            sent = True
    elif _is_meeting(ev) and _is_organizer(ev):
        notes.append(
            "Saved in your calendar only. Attendees were not notified; pass "
            "send_update=true to send them the update."
        )
    out = {"status": "updated", "entry_id": entry_id, "update_sent": sent}
    if removed:
        out["removed_attendees"] = removed
    if notes:
        out["notes"] = notes
    return out


def delete_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    send_cancellation: bool = True,
    occurrence_start: str | None = None,
    apply_to_series: bool = False,
) -> dict[str, Any]:
    ev = _target_event(namespace, entry_id, occurrence_start)
    _guard_series(ev, apply_to_series, "delete")
    subject = _safe_get(ev, "Subject", "")
    notes: list[str] = []
    cancelled = False
    if _is_meeting(ev) and _is_organizer(ev) and send_cancellation:
        if draft_only():
            raise OutlookError(
                "This meeting has attendees and draft-only mode is on, so a "
                "cancellation can't be sent. Cancel it from Outlook, or pass "
                "send_cancellation=false to delete it from your calendar only."
            )
        ev.MeetingStatus = OL_MEETING_CANCELED
        ev.Save()
        ev.Send()
        cancelled = True
    elif _is_meeting(ev) and not _is_organizer(ev):
        notes.append(
            "You are an attendee: the meeting was removed from your calendar "
            "without replying. Use outlook_respond_event to decline instead "
            "if the organizer should know."
        )
    ev.Delete()
    out = {
        "status": "deleted",
        "subject": subject,
        "entry_id": entry_id,
        "cancellation_sent": cancelled,
    }
    if notes:
        out["notes"] = notes
    return out


def respond_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    response: str,
    send_response: bool = True,
    comment: str | None = None,
    occurrence_start: str | None = None,
) -> dict[str, Any]:
    ev = _target_event(namespace, entry_id, occurrence_start)
    code = {
        "accept": OL_MEETING_ACCEPTED,
        "tentative": OL_MEETING_TENTATIVE,
        "decline": OL_MEETING_DECLINED,
    }.get(response.lower())
    if code is None:
        raise OutlookError("response must be one of: 'accept', 'tentative', 'decline'.")
    resp = ev.Respond(code, True)
    status = "responded"
    notes: list[str] = []
    if resp is not None:
        if comment:
            resp.Body = comment
        if send_response and not draft_only():
            resp.Send()
        elif send_response:
            resp.Save()
            status = "response_saved_to_drafts"
            notes.append(DRAFT_ONLY_NOTE)
    out = {"status": status, "response": response}
    if notes:
        out["notes"] = notes
    return out
