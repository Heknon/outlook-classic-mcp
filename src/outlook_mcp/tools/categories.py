"""MCP tool wrappers for Outlook categories (colour-coded labels)."""

from __future__ import annotations

from typing import Annotated, Optional, Union

from pydantic import Field

from outlook_mcp.client import categories as cat_client
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call

COLOR_DESC = f"Colour name: {cat_client.COLOR_CHOICES} (also grey, navy, gold, brown, cranberry, plum)."

_WRITE = {
    "readOnlyHint": False,
    "destructiveHint": False,
    "idempotentHint": True,
    "openWorldHint": False,
}


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_categories",
        annotations={
            "title": "List Outlook color categories",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_categories(
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List the colour categories in this Outlook profile, with their colours."""
        data = await bridge.call(cat_client.list_categories)
        return format_response(data, response_format)

    @mcp.tool(name="outlook_create_category", annotations={"title": "Create a colour category", **_WRITE})
    @safe_call
    async def outlook_create_category(
        name: Annotated[str, Field(min_length=1, description="Category name, e.g. 'Customer A'.")],
        color: Annotated[str, Field(description=COLOR_DESC)] = "blue",
    ) -> str:
        """Create a colour category (for mail, meetings and tasks)."""
        data = await bridge.call(cat_client.create_category, name=name, color=color)
        return format_response(data, "json")

    @mcp.tool(name="outlook_update_category", annotations={"title": "Recolour or rename a category", **_WRITE})
    @safe_call
    async def outlook_update_category(
        name: Annotated[str, Field(min_length=1, description="Existing category name.")],
        color: Annotated[Optional[str], Field(description=COLOR_DESC)] = None,
        new_name: Annotated[Optional[str], Field(description="Rename to this.")] = None,
    ) -> str:
        """Change a category's colour and/or name."""
        data = await bridge.call(cat_client.update_category, name=name, color=color, new_name=new_name)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_category",
        annotations={"title": "Delete a category", **_WRITE, "destructiveHint": True},
    )
    @safe_call
    async def outlook_delete_category(
        name: Annotated[str, Field(min_length=1, description="Category to remove from the list.")],
    ) -> str:
        """Remove a category from the master list (tagged items keep the text, without colour)."""
        data = await bridge.call(cat_client.delete_category, name=name)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_set_category",
        annotations={
            "title": "Categorize mails / meetings / tasks",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_set_category(
        categories: Annotated[
            Union[str, list[str]],
            Field(
                description=(
                    "Category name(s): a list, or comma-separated ('Work, Follow-up'). "
                    "With mode='replace', an empty string clears all categories."
                ),
            ),
        ],
        entry_id: Annotated[Optional[str], Field(description="EntryID of one item (mail, meeting or task).")] = None,
        entry_ids: Annotated[
            Optional[list[str]],
            Field(max_length=500, description="Several items at once, e.g. the entry_ids from a search."),
        ] = None,
        mode: Annotated[
            str,
            Field(description="'replace' (default) sets exactly these, 'add' keeps existing ones, 'remove' takes these off."),
        ] = "replace",
        occurrence_start: Annotated[
            Optional[str],
            Field(description="For one occurrence of a recurring meeting: its 'start' from outlook_list_events."),
        ] = None,
        create_missing: Annotated[
            bool,
            Field(description="Create categories that don't exist yet (so they show in colour)."),
        ] = True,
        color: Annotated[Optional[str], Field(description="Colour for newly created categories. " + COLOR_DESC)] = None,
    ) -> str:
        """Colour-code mails, meetings and tasks with categories.

        Works on one item or many (bulk). Meetings take their colour in the
        calendar from their first category. New category names are created
        automatically (in `color`, default blue).
        """
        data = await bridge.call(
            cat_client.set_category,
            entry_id=entry_id,
            entry_ids=entry_ids,
            categories=categories,
            mode=mode,
            occurrence_start=occurrence_start,
            create_missing=create_missing,
            color=color,
        )
        return format_response(data, "json")
