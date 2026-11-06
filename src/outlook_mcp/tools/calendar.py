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
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List calendar events in a date range, including recurring instances.

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
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Fetch an event with attendees, body, and recurrence info."""
        data = await bridge.call(
            cal_client.get_event, entry_id=entry_id, occurrence_start=occurrence_start
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
