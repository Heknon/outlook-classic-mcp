"""MCP tool wrappers for calendar / meetings."""

from __future__ import annotations

from typing import Annotated, Optional

from mcp.types import CallToolResult
from pydantic import Field

from outlook_mcp.client import calendar as cal_client
from outlook_mcp.schemas import Recurrence
from outlook_mcp.ui import ui_meta, ui_result
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


OCCURRENCE_DESC = (
    "For one occurrence of a recurring meeting: that occurrence's 'start' "
    "exactly as outlook_list_events returned it (the entry_id of an "
    "occurrence is the series id)."
)

SERIES_DESC = (
    "Required to change or delete a whole recurring series (when no "
    "occurrence_start is given). Guards against moving every occurrence "
    "by accident."
)


MAILBOX_DESC = (
    "Someone else's calendar: their email address or name. Needs read access "
    "to their calendar (they shared it, or you're a delegate); otherwise use "
    "outlook_get_free_busy, which works for anyone in the organization."
)

READ_ONLY = {
    "readOnlyHint": True,
    "destructiveHint": False,
    "idempotentHint": True,
    "openWorldHint": False,
}


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_events",
        annotations={
            "title": "List Outlook calendar events",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("calendar"),
        structured_output=False,
    )
    @safe_call
    async def outlook_list_events(
        start: Annotated[Optional[str], Field(description="ISO-8601 start of range. Defaults to now.")] = None,
        end: Annotated[Optional[str], Field(description="ISO-8601 end of range. Defaults to start+14d.")] = None,
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        include_recurrences: Annotated[bool, Field()] = True,
        mailbox: Annotated[Optional[str], Field(description=MAILBOX_DESC)] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List calendar events in a date range — yours, or a colleague's with `mailbox`.

        Filters by overlap with [start, end), so multi-day/all-day events that
        started earlier are included. Response includes "truncated": true if
        more matching events exist beyond `limit`.
        """
        data = await bridge.call(
            cal_client.list_events,
            start=start,
            end=end,
            limit=limit,
            include_recurrences=include_recurrences,
            mailbox=mailbox,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_get_event",
        annotations={
            "title": "Get Outlook calendar event",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        structured_output=False,
    )
    @safe_call
    async def outlook_get_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
        occurrence_start: Annotated[Optional[str], Field(description=OCCURRENCE_DESC)] = None,
        mailbox: Annotated[
            Optional[str],
            Field(description="The calendar owner, when the event came from outlook_list_events with a mailbox."),
        ] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Fetch an event with attendees, body, and recurrence info."""
        data = await bridge.call(
            cal_client.get_event,
            entry_id=entry_id,
            occurrence_start=occurrence_start,
            mailbox=mailbox,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_create_event",
        annotations={
            "title": "Create Outlook calendar event or meeting",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_create_event(
        subject: Annotated[str, Field(description="Event subject.")],
        start: Annotated[
            str,
            Field(description="ISO-8601 local start (e.g. '2026-05-01T14:00:00'), or a date for all-day events."),
        ],
        end: Annotated[
            Optional[str],
            Field(description="ISO-8601 end. Required unless all_day=true (then the last day, inclusive; defaults to one day)."),
        ] = None,
        location: Annotated[Optional[str], Field()] = None,
        body: Annotated[Optional[str], Field()] = None,
        attendees: Annotated[
            Optional[list[str]],
            Field(description="Required attendees. Adding any attendee turns this into a meeting."),
        ] = None,
        optional_attendees: Annotated[Optional[list[str]], Field(description="Optional attendees.")] = None,
        is_online_meeting: Annotated[bool, Field(description="Reserved; ignored on classic Outlook COM.")] = False,
        reminder_minutes: Annotated[Optional[int], Field(ge=0, le=10080, description="Null for no reminder.")] = 15,
        recurrence: Annotated[
            Optional[Recurrence],
            Field(description="Recurrence pattern. Omit for a single occurrence."),
        ] = None,
        all_day: Annotated[bool, Field(description="Create an all-day event.")] = False,
        send_invites: Annotated[
            bool,
            Field(
                description=(
                    "For meetings: send the invitations now. False saves the meeting "
                    "in the calendar unsent, for the user to review and send from Outlook."
                )
            ),
        ] = True,
        busy_status: Annotated[
            Optional[str],
            Field(description="free, tentative, busy, out_of_office or working_elsewhere."),
        ] = None,
    ) -> str:
        """Create a calendar event or meeting.

        A meeting (any attendees) sends invitations as soon as this succeeds
        unless send_invites=false — confirm attendees and times first.
        """
        data = await bridge.call(
            cal_client.create_event,
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            attendees=attendees,
            optional_attendees=optional_attendees,
            is_online_meeting=is_online_meeting,
            reminder_minutes=reminder_minutes,
            recurrence=recurrence,
            all_day=all_day,
            send_invites=send_invites,
            busy_status=busy_status,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_update_event",
        annotations={
            "title": "Update Outlook calendar event",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_update_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
        subject: Annotated[Optional[str], Field()] = None,
        start: Annotated[Optional[str], Field(description="New ISO-8601 local start.")] = None,
        end: Annotated[Optional[str], Field(description="New ISO-8601 local end.")] = None,
        location: Annotated[Optional[str], Field()] = None,
        body: Annotated[Optional[str], Field()] = None,
        all_day: Annotated[Optional[bool], Field(description="Turn all-day on or off.")] = None,
        reminder_minutes: Annotated[Optional[int], Field(ge=0, le=10080)] = None,
        busy_status: Annotated[Optional[str], Field()] = None,
        add_attendees: Annotated[Optional[list[str]], Field(description="Required attendees to add.")] = None,
        add_optional_attendees: Annotated[Optional[list[str]], Field(description="Optional attendees to add.")] = None,
        remove_attendees: Annotated[Optional[list[str]], Field(description="Attendees to remove (address or name).")] = None,
        send_update: Annotated[
            bool,
            Field(
                description=(
                    "For meetings you organize: send the updated meeting to attendees. "
                    "False only changes your calendar."
                )
            ),
        ] = False,
        occurrence_start: Annotated[Optional[str], Field(description=OCCURRENCE_DESC)] = None,
        apply_to_series: Annotated[bool, Field(description=SERIES_DESC)] = False,
    ) -> str:
        """Update an event. Only non-null fields are written."""
        data = await bridge.call(
            cal_client.update_event,
            entry_id=entry_id,
            subject=subject,
            start=start,
            end=end,
            location=location,
            body=body,
            all_day=all_day,
            reminder_minutes=reminder_minutes,
            busy_status=busy_status,
            add_attendees=add_attendees,
            add_optional_attendees=add_optional_attendees,
            remove_attendees=remove_attendees,
            send_update=send_update,
            occurrence_start=occurrence_start,
            apply_to_series=apply_to_series,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_event",
        annotations={
            "title": "Delete or cancel Outlook calendar event",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_delete_event(
        entry_id: Annotated[str, Field(description="EntryID of the event.")],
        send_cancellation: Annotated[
            bool,
            Field(description="For meetings you organize: send a cancellation to attendees before deleting."),
        ] = True,
        occurrence_start: Annotated[Optional[str], Field(description=OCCURRENCE_DESC)] = None,
        apply_to_series: Annotated[bool, Field(description=SERIES_DESC)] = False,
    ) -> str:
        """Delete an event; for meetings you organize, cancel it for attendees too."""
        data = await bridge.call(
            cal_client.delete_event,
            entry_id=entry_id,
            send_cancellation=send_cancellation,
            occurrence_start=occurrence_start,
            apply_to_series=apply_to_series,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_respond_event",
        annotations={
            "title": "Respond to a meeting invite",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_respond_event(
        entry_id: Annotated[str, Field(description="EntryID of the meeting.")],
        response: Annotated[str, Field(description="One of: 'accept', 'tentative', 'decline'.")],
        send_response: Annotated[bool, Field(description="Send the response to the organizer.")] = True,
        comment: Annotated[Optional[str], Field(description="Optional note to include in the response.")] = None,
        occurrence_start: Annotated[Optional[str], Field(description=OCCURRENCE_DESC)] = None,
    ) -> str:
        """Accept / tentatively accept / decline a meeting invite."""
        data = await bridge.call(
            cal_client.respond_event,
            entry_id=entry_id,
            response=response,
            send_response=send_response,
            comment=comment,
            occurrence_start=occurrence_start,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_get_free_busy",
        annotations={"title": "Get free/busy for people", **READ_ONLY},
    )
    @safe_call
    async def outlook_get_free_busy(
        people: Annotated[
            list[str],
            Field(min_length=1, max_length=30, description="Email addresses or names (resolved in the address book)."),
        ],
        start: Annotated[Optional[str], Field(description="ISO-8601 local start. Defaults to now.")] = None,
        end: Annotated[Optional[str], Field(description="ISO-8601 local end. Defaults to start + 1 day.")] = None,
        interval_minutes: Annotated[int, Field(ge=5, le=60, description="Resolution of the free/busy data.")] = 30,
        include_me: Annotated[bool, Field(description="Include your own free/busy too.")] = False,
    ) -> str:
        """When people are busy, from Exchange free/busy.

        Works for anyone in the organization without access to their
        calendar (no meeting details, only busy / tentative / out of office).
        """
        data = await bridge.call(
            cal_client.get_free_busy,
            people=people,
            start=start,
            end=end,
            interval_minutes=interval_minutes,
            include_me=include_me,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_find_meeting_times",
        annotations={"title": "Find meeting times (scheduling assistant)", **READ_ONLY},
    )
    @safe_call
    async def outlook_find_meeting_times(
        attendees: Annotated[
            list[str],
            Field(min_length=1, max_length=30, description="Required attendees (addresses or names)."),
        ],
        duration_minutes: Annotated[int, Field(ge=5, le=600)] = 30,
        optional_attendees: Annotated[
            Optional[list[str]],
            Field(description="Optional attendees: reported per slot when busy, never exclude a slot."),
        ] = None,
        start: Annotated[Optional[str], Field(description="Search from (ISO-8601 local). Defaults to now.")] = None,
        end: Annotated[Optional[str], Field(description="Search until. Defaults to start + 7 days.")] = None,
        day_start: Annotated[str, Field(description="Working day start, HH:MM local.")] = "09:00",
        day_end: Annotated[str, Field(description="Working day end, HH:MM local.")] = "17:00",
        working_days: Annotated[
            Optional[list[str]],
            Field(
                description=(
                    "Days to consider, e.g. ['sun','mon','tue','wed','thu'] for a "
                    "Sunday–Thursday week. Defaults to Monday–Friday."
                )
            ),
        ] = None,
        step_minutes: Annotated[
            int, Field(description="Start times are aligned to this (15, 30 or 60).")
        ] = 30,
        include_me: Annotated[bool, Field(description="Require your own calendar to be free too.")] = True,
        tentative_is_free: Annotated[bool, Field(description="Treat tentative as available.")] = False,
        max_results: Annotated[int, Field(ge=1, le=50)] = 10,
        max_per_day: Annotated[int, Field(ge=1, le=20, description="Spread results across days.")] = 3,
    ) -> str:
        """Find times when all required attendees (and you) are free.

        Uses Exchange free/busy, like Outlook's Scheduling Assistant. Pick a
        slot with the user, then create the meeting with outlook_create_event.
        """
        if step_minutes not in (5, 10, 15, 20, 30, 60):
            raise ValueError("step_minutes must be one of 5, 10, 15, 20, 30, 60.")
        data = await bridge.call(
            cal_client.find_meeting_times,
            attendees=attendees,
            duration_minutes=duration_minutes,
            optional_attendees=optional_attendees,
            start=start,
            end=end,
            day_start=day_start,
            day_end=day_end,
            working_days=working_days,
            step_minutes=step_minutes,
            include_me=include_me,
            tentative_is_free=tentative_is_free,
            max_results=max_results,
            max_per_day=max_per_day,
        )
        return format_response(data, "json")
