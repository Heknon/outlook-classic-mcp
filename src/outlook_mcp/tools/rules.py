"""MCP tool wrappers for mail rules."""

from __future__ import annotations

from typing import Annotated, Optional

from pydantic import Field

from outlook_mcp.client import rules as rules_client
from outlook_mcp.schemas import RuleActions, RuleConditions
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call

CONFIRM_DESC = (
    "False (default) returns a preview and changes nothing. Show the preview "
    "to the user and call again with true only after they agree."
)


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_rules",
        annotations={
            "title": "List Outlook mail rules",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_rules(
        details: Annotated[
            bool,
            Field(description="Also explain each rule: its conditions, exceptions and actions."),
        ] = False,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> str:
        """List mail rules with their enabled state (and what they do, with details=true)."""
        data = await bridge.call(rules_client.list_rules, details=details)
        return format_response(data, "json" if details else response_format)

    @mcp.tool(
        name="outlook_toggle_rule",
        annotations={
            "title": "Enable / disable an Outlook rule",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_toggle_rule(
        rule_name: Annotated[str, Field(min_length=1, description="Rule name (case-insensitive).")],
        enabled: Annotated[bool, Field(description="True to enable, False to disable.")],
    ) -> str:
        """Enable or disable a mail rule. This changes live mail handling — confirm first."""
        data = await bridge.call(rules_client.toggle_rule, rule_name=rule_name, enabled=enabled)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_create_rule",
        annotations={
            "title": "Create a mail rule (preview first)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_create_rule(
        name: Annotated[str, Field(min_length=1, description="Rule name.")],
        conditions: Annotated[RuleConditions, Field(description="When the rule applies. Empty = every incoming message.")],
        actions: Annotated[RuleActions, Field(description="What it does. At least one action.")],
        exceptions: Annotated[
            Optional[RuleConditions], Field(description="Except when any of these hold.")
        ] = None,
        enabled: Annotated[bool, Field()] = True,
        replace: Annotated[bool, Field(description="Replace an existing rule with the same name.")] = False,
        run_now: Annotated[bool, Field(description="After creating, also apply it to mail already in the Inbox.")] = False,
        confirm: Annotated[bool, Field(description=CONFIRM_DESC)] = False,
    ) -> str:
        """Create an incoming-mail rule, like Outlook's Rules wizard.

        Two steps: the first call (confirm=false) returns the rule in plain
        terms and the recent Inbox mail it would have matched; create it
        with confirm=true once the user agrees.

        Forward/redirect actions send mail automatically: they're refused in
        draft-only mode, and otherwise only allowed to addresses in your own
        organization unless OUTLOOK_MCP_ALLOW_EXTERNAL_FORWARD=1.
        """
        data = await bridge.call(
            rules_client.create_rule,
            name=name,
            conditions=conditions,
            actions=actions,
            exceptions=exceptions,
            enabled=enabled,
            replace=replace,
            run_now=run_now,
            confirm=confirm,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_rule",
        annotations={
            "title": "Delete a mail rule (preview first)",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_rule(
        name: Annotated[str, Field(min_length=1, description="Rule name.")],
        confirm: Annotated[bool, Field(description=CONFIRM_DESC)] = False,
    ) -> str:
        """Delete a mail rule. The first call shows the rule; confirm=true deletes it."""
        data = await bridge.call(rules_client.delete_rule, name=name, confirm=confirm)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_run_rule",
        annotations={
            "title": "Run a rule on existing mail (preview first)",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_run_rule(
        name: Annotated[str, Field(min_length=1, description="Rule name.")],
        folder: Annotated[str, Field(description="Folder to run it on.")] = "inbox",
        include_subfolders: Annotated[bool, Field()] = False,
        which: Annotated[str, Field(description="'all', 'read' or 'unread' messages.")] = "all",
        confirm: Annotated[bool, Field(description=CONFIRM_DESC)] = False,
    ) -> str:
        """Apply a rule to mail already in a folder (Outlook's "Run Rules Now").

        The first call previews which messages it would affect; running it
        moves/deletes/categorizes them for real, so confirm with the user.
        """
        data = await bridge.call(
            rules_client.run_rule,
            name=name,
            folder=folder,
            include_subfolders=include_subfolders,
            which=which,
            confirm=confirm,
        )
        return format_response(data, "json")
