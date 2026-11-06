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


def list_events(
    outlook: Any,
    namespace: Any,
    *,
    start: str | None = None,
    end: str | None = None,
    limit: int = 50,
    include_recurrences: bool = True,
) -> dict[str, Any]:
    cal = namespace.GetDefaultFolder(OL_FOLDER_CALENDAR)
    items = cal.Items
    # Outlook requires the collection sorted ascending by [Start] BEFORE
    # IncludeRecurrences is set, or recurring series expand incorrectly.
    items.Sort("[Start]")
    if include_recurrences:
        items.IncludeRecurrences = True

    start_dt = from_iso(start) or dt.datetime.now()
    end_dt = from_iso(end) or (start_dt + dt.timedelta(days=14))

    # DASL property query (locale-independent ISO dates) instead of a plain
    # Jet literal — see to_dasl_date() for why the plain "[Start] >= 'mm/dd
    # /yyyy ...'" form silently returns the wrong events on non-US locales.
    restrict = (
        f"@SQL=\"urn:schemas:calendar:dtstart\" >= "
        f"'{to_dasl_date(start_dt - _OVERLAP_LOOKBACK)}' AND "
        f"\"urn:schemas:calendar:dtstart\" <= '{to_dasl_date(end_dt)}'"
    )
    filtered = items.Restrict(restrict)

    results: list[dict[str, Any]] = []
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
        if len(results) >= limit:
            truncated = True
            break
        results.append(_event_summary(ev))

    return {
        "start": start_dt.isoformat(),
        "end": end_dt.isoformat(),
        "count": len(results),
        "items": results,
        "truncated": truncated,
    }


def get_event(
    outlook: Any,
    namespace: Any,
    *,
    entry_id: str,
    occurrence_start: str | None = None,
) -> dict[str, Any]:
    return _event_full(_target_event(namespace, entry_id, occurrence_start))


def _target_event(namespace: Any, entry_id: str, occurrence_start: str | None) -> Any:
    """Resolve an event, or one occurrence of a recurring series.

    ``list_events`` returns the *series* EntryID for every occurrence of a
    recurring meeting, so acting on that id alone would hit the whole
    series. ``occurrence_start`` (the occurrence's original start, as
    listed) selects the single occurrence instead.
    """
    ev = get_item_by_id(namespace, entry_id)
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
