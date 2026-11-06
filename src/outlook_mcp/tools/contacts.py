"""MCP tool wrappers for contacts."""

from __future__ import annotations

from typing import Annotated

from mcp.types import CallToolResult
from pydantic import Field

from outlook_mcp.client import contacts as contacts_client
from outlook_mcp.ui import ui_meta, ui_result
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_contacts",
        annotations={
            "title": "List Outlook contacts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("contacts"),
        structured_output=False,
    )
    @safe_call
    async def outlook_list_contacts(
        limit: Annotated[int, Field(ge=1, le=200)] = 50,
        offset: Annotated[int, Field(ge=0)] = 0,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List saved contacts from every contact folder in every store.

        Note: on corporate accounts the personal contact folders are often
        nearly empty — colleagues live in the directory instead. Use
        outlook_search_contacts (which also searches the directory) to
        find people.
        """
        data = await bridge.call(
            contacts_client.list_contacts, limit=limit, offset=offset
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_search_contacts",
        annotations={
            "title": "Search Outlook contacts",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("contacts"),
        structured_output=False,
    )
    @safe_call
    async def outlook_search_contacts(
        query: Annotated[
            str,
            Field(
                min_length=1,
                description="Search words; all must match (name, email, company, or title).",
            ),
        ],
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        include_directory: Annotated[
            bool,
            Field(
                description=(
                    "Also search the org directory (Exchange Global Address "
                    "List). Can take a few seconds on large directories."
                ),
            ),
        ] = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Search saved contacts and (by default) the org directory."""
        data = await bridge.call(
            contacts_client.search_contacts,
            query=query,
            limit=limit,
            include_directory=include_directory,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_resolve_name",
        annotations={
            "title": "Check names (Ctrl+K): find a person's email address",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_resolve_name(
        name: Annotated[
            str,
            Field(
                min_length=1,
                description="Display name, partial name, alias, or address (like Ctrl+K in Outlook).",
            ),
        ],
        max_candidates: Annotated[int, Field(ge=1, le=50)] = 10,
    ) -> str:
        """Check a name like Ctrl+K: get someone's email address from their name.

        Returns status 'resolved' with the person (address, title,
        department, office, phone), 'ambiguous' with the candidates to
        choose from (ask the user which one), or 'not_found'. Searches your
        contacts and the organization's address book.
        """
        data = await bridge.call(
            contacts_client.resolve_name, name=name, max_candidates=max_candidates
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_get_contact",
        annotations={
            "title": "Get Outlook contact",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_get_contact(
        entry_id: Annotated[str, Field(description="EntryID of the contact.")],
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """Fetch full details for one contact."""
        data = await bridge.call(contacts_client.get_contact, entry_id=entry_id)
        return format_response(data, response_format)
