"""MCP tool wrappers for inbox-zero triage: Quick Steps, bulk filing, Waiting."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import quicksteps as qs_client
from outlook_mcp.client import triage as triage_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call

_READ = {"readOnlyHint": True, "destructiveHint": False, "idempotentHint": True, "openWorldHint": False}
_WRITE = {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": False}


def register(mcp, bridge) -> None:
    @mcp.tool(name="outlook_list_quick_steps", annotations={"title": "List Outlook Quick Steps", **_READ})
    @safe_call
    async def outlook_list_quick_steps() -> str:
        """The user's Outlook Quick Steps and what each one does.

        Use them as the user's filing buckets (e.g. Archive / Action /
        Waiting / Read). `runnable` says whether outlook_run_quick_step can
        carry it out; `problems` explains why not.
        """
        data = await bridge.call(qs_client.list_quick_steps)
        return format_response(data, "json")

    @mcp.tool(name="outlook_run_quick_step", annotations={"title": "Run a Quick Step on messages", **_WRITE})
    @safe_call
    async def outlook_run_quick_step(
        name: Annotated[str, Field(min_length=1, description="Quick Step name (as listed).")],
        entry_ids: Annotated[list[str], Field(min_length=1, max_length=500, description="Messages to apply it to.")],
        confirm: Annotated[
            bool, Field(description="Needed only for Quick Steps that delete; otherwise ignored.")
        ] = False,
    ) -> str:
        """Do what a Quick Step does (move, copy, mark read, categorize, flag ...)
        to the given messages, like clicking it in Outlook.

        Quick Steps that reply, forward or create items aren't run. Ones that
        delete need confirm=true.
        """
        data = await bridge.call(qs_client.run_quick_step, name=name, entry_ids=entry_ids, confirm=confirm)
        return format_response(data, "json")

    @mcp.tool(name="outlook_move_mails", annotations={"title": "Move several messages to a folder", **_WRITE})
    @safe_call
    async def outlook_move_mails(
        entry_ids: Annotated[list[str], Field(min_length=1, max_length=500)],
        target_folder: Annotated[str, Field(description="Destination, e.g. 'Inbox/1-Action' or 'archive'.")],
        mark_read: Annotated[Optional[bool], Field(description="Also mark read (true) / unread (false).")] = None,
        add_categories: Annotated[Optional[list[str]], Field(description="Also add these categories.")] = None,
        flag: Annotated[Optional[bool], Field(description="Also flag (true) / unflag (false).")] = None,
    ) -> str:
        """File many messages at once. Returns each message's new entry_id."""
        data = await bridge.call(
            triage_client.move_mails,
            entry_ids=entry_ids,
            target_folder=target_folder,
            mark_read=mark_read,
            add_categories=add_categories,
            flag=flag,
        )
        return format_response(data, "json")

    @mcp.tool(name="outlook_check_waiting", annotations={"title": "Which Waiting items got replies", **_READ})
    @safe_call
    async def outlook_check_waiting(
        folder: Annotated[
            Optional[str],
            Field(description="The Waiting folder. Default: the folder the user's 'Waiting' Quick Step moves to, else a folder named like 'Waiting'."),
        ] = None,
        move_answered_to: Annotated[
            Optional[str],
            Field(description="Move messages that got a reply to this folder (e.g. 'inbox' or 'Inbox/1-Action'). Default: report only."),
        ] = None,
        limit: Annotated[int, Field(ge=1, le=1000)] = 200,
    ) -> str:
        """For each message in the Waiting folder, check its thread for a reply
        from someone else since it was filed.

        Returns answered ones (with the replies) and still-waiting ones
        (longest waiting first, with days waiting).
        """
        data = await bridge.call(
            triage_client.check_waiting, folder=folder, move_answered_to=move_answered_to, limit=limit
        )
        return format_response(data, "json")
