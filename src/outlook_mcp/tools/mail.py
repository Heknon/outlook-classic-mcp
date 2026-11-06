"""MCP tool wrappers for mail."""

from __future__ import annotations

from typing import Annotated, Optional, Union

from mcp.types import CallToolResult
from pydantic import Field

from outlook_mcp.client import mail as mail_client
from outlook_mcp.ui import ui_meta, ui_result
from outlook_mcp.utils.formatting import format_response
from outlook_mcp.utils.safety import safe_call


SAVE_ONLY_DESC = (
    "If true, save to Drafts instead of sending and return the draft's "
    "entry_id. Use this whenever the user asked for a draft or hasn't "
    "explicitly approved sending."
)

SIGNATURE_DESC = (
    "Signature to add below your text: true or 'default' for the account's "
    "default Outlook signature (the reply/forward default for replies), or "
    "a signature name from outlook_list_signatures. Omit for none."
)

SignatureArg = Optional[Union[bool, str]]


def register(mcp, bridge) -> None:
    @mcp.tool(
        name="outlook_list_mails",
        annotations={
            "title": "List Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_list_mails(
        folder: Annotated[
            str,
            Field(
                description=(
                    "Folder name. Either a well-known name (inbox, sent, drafts, "
                    "deleted, junk) or a path like 'Inbox/Projects/Quinn'."
                ),
            ),
        ] = "inbox",
        limit: Annotated[int, Field(ge=1, le=100, description="Max items.")] = 25,
        offset: Annotated[int, Field(ge=0, description="Pagination offset.")] = 0,
        unread_only: Annotated[bool, Field(description="Return only unread.")] = False,
        since: Annotated[Optional[str], Field(description="ISO-8601 lower bound on ReceivedTime.")] = None,
        until: Annotated[Optional[str], Field(description="ISO-8601 upper bound on ReceivedTime.")] = None,
        from_address: Annotated[Optional[str], Field(description="Substring match on sender email.")] = None,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """List mail items from a folder, newest first."""
        data = await bridge.call(
            mail_client.list_mails,
            folder=folder,
            limit=limit,
            offset=offset,
            unread_only=unread_only,
            since=since,
            until=until,
            from_address=from_address,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_search_mails",
        annotations={
            "title": "Search Outlook mails",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_search_mails(
        query: Annotated[str, Field(min_length=1, description="Search keywords or DASL filter.")],
        folder: Annotated[str, Field(description="Folder to search in.")] = "inbox",
        scope: Annotated[
            str,
            Field(
                description=(
                    "Where to look: 'subject_body' (default), 'subject', 'from', "
                    "or 'dasl' to pass `query` as a raw DASL @SQL filter."
                ),
            ),
        ] = "subject_body",
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Search a mail folder by subject, body, or sender."""
        data = await bridge.call(
            mail_client.search_mails,
            query=query,
            folder=folder,
            limit=limit,
            scope=scope,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_search_mails_by_person",
        annotations={
            "title": "Find mails from/to a person",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-list"),
        structured_output=False,
    )
    @safe_call
    async def outlook_search_mails_by_person(
        person: Annotated[
            str,
            Field(min_length=1, description="Name, partial name, alias or email address (resolved like Ctrl+K)."),
        ],
        direction: Annotated[
            str,
            Field(description="'from' (sent by them), 'to' (they were on To/CC), or 'any'."),
        ] = "any",
        folders: Annotated[
            Optional[list[str]],
            Field(description="Folders to search. Default: inbox and sent (inbox only for 'from')."),
        ] = None,
        since: Annotated[Optional[str], Field(description="Only mails received after this ISO-8601 time.")] = None,
        limit: Annotated[int, Field(ge=1, le=100)] = 25,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Find mails exchanged with a person, given just their name.

        The name is resolved against your contacts and the address book first.
        If several people match, no search is run: the result has
        status 'ambiguous' and the candidates — ask the user which one, then
        call again with that person's email address.
        """
        data = await bridge.call(
            mail_client.search_mails_by_person,
            person=person,
            direction=direction,
            folders=folders,
            since=since,
            limit=limit,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_get_mail",
        annotations={
            "title": "Get full Outlook mail",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
        meta=ui_meta("mail-view"),
        structured_output=False,
    )
    @safe_call
    async def outlook_get_mail(
        entry_id: Annotated[str, Field(min_length=1, description="EntryID of the mail item.")],
        include_body: Annotated[bool, Field(description="Include the plain-text body.")] = True,
        include_html: Annotated[
            bool,
            Field(
                description=(
                    "Also include the raw HTML body. Off by default — it is "
                    "usually huge and rarely needed; the plain-text body "
                    "carries the same content."
                ),
            ),
        ] = False,
        max_body_chars: Annotated[
            int,
            Field(
                ge=0,
                description="Truncate the body beyond this many chars (0 = no limit).",
            ),
        ] = 10000,
        include_attachments: Annotated[
            bool,
            Field(
                description=(
                    "Include the attachment list (filename/size per attachment). "
                    "Off for bulk reads that don't need attachments — it's the "
                    "slowest and most crash-prone part of fetching a mail. When "
                    "False, 'attachments' is null and Attachments isn't touched "
                    "at all (no attachment_count either)."
                ),
            ),
        ] = True,
        response_format: Annotated[str, Field(description="'markdown' or 'json'.")] = "markdown",
    ) -> CallToolResult:
        """Fetch body, headers, and attachment list for one mail item.

        If the response has body_truncated=true, re-call with a higher
        max_body_chars to read more.
        """
        data = await bridge.call(
            mail_client.get_mail,
            entry_id=entry_id,
            include_body=include_body,
            include_html=include_html,
            max_body_chars=max_body_chars,
            include_attachments=include_attachments,
        )
        return ui_result(format_response(data, response_format), data)

    @mcp.tool(
        name="outlook_send_mail",
        annotations={
            "title": "Send Outlook mail (or save as draft)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_send_mail(
        to: Annotated[list[str], Field(min_length=1, description="Recipient addresses.")],
        subject: Annotated[str, Field(description="Subject line.")],
        body: Annotated[str, Field(description="Message body. Plain text unless html=True.")],
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as HTML.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths to local files.")] = None,
        importance: Annotated[str, Field(description="One of: 'low', 'normal', 'high'.")] = "normal",
        save_only: Annotated[bool, Field(description=SAVE_ONLY_DESC)] = False,
        account: Annotated[
            Optional[str],
            Field(
                description=(
                    "Optional sender account (SMTP address or display name, e.g. "
                    "'info@example.com'). The mail is sent from this account and, "
                    "with save_only=True, saved to this account's Drafts folder. "
                    "Defaults to the profile's default account."
                )
            ),
        ] = None,
        signature: Annotated[SignatureArg, Field(description=SIGNATURE_DESC)] = None,
    ) -> str:
        """Compose a new mail and send it, or save it to Drafts with save_only=True."""
        data = await bridge.call(
            mail_client.send_mail,
            to=to,
            subject=subject,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            importance=importance,
            save_only=save_only,
            account=account,
            signature=signature,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_reply_mail",
        annotations={
            "title": "Reply to Outlook mail (or save reply as draft)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_reply_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to reply to.")],
        body: Annotated[
            str,
            Field(
                description=(
                    "Your reply text. It is placed above the quoted original, "
                    "which Outlook keeps intact. Plain text unless html=True."
                )
            ),
        ],
        reply_all: Annotated[bool, Field(description="Reply to all recipients.")] = False,
        html: Annotated[bool, Field(description="Treat body as an HTML fragment.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths of files to attach.")] = None,
        to: Annotated[Optional[list[str]], Field(description="Extra To recipients, added to the reply's own.")] = None,
        cc: Annotated[Optional[list[str]], Field(description="Extra CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        subject: Annotated[Optional[str], Field(description="Replace the subject (default keeps 'RE: ...').")] = None,
        importance: Annotated[Optional[str], Field(description="One of: 'low', 'normal', 'high'.")] = None,
        signature: Annotated[SignatureArg, Field(description=SIGNATURE_DESC)] = None,
        save_only: Annotated[bool, Field(description=SAVE_ONLY_DESC)] = False,
    ) -> str:
        """Reply (or reply-all) to a mail, keeping the thread and quoted history.

        With save_only=True the reply is saved to Drafts (still threaded) and
        its entry_id returned, so the user can review it before sending.
        """
        data = await bridge.call(
            mail_client.reply_mail,
            entry_id=entry_id,
            body=body,
            reply_all=reply_all,
            html=html,
            attachments=attachments,
            to=to,
            cc=cc,
            bcc=bcc,
            subject=subject,
            importance=importance,
            signature=signature,
            save_only=save_only,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_forward_mail",
        annotations={
            "title": "Forward Outlook mail (or save forward as draft)",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_forward_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to forward.")],
        to: Annotated[
            Optional[list[str]],
            Field(description="Forward recipients. Required unless save_only=True."),
        ] = None,
        body: Annotated[str, Field(description="Optional note above the forwarded mail.")] = "",
        cc: Annotated[Optional[list[str]], Field(description="CC recipients.")] = None,
        bcc: Annotated[Optional[list[str]], Field(description="BCC recipients.")] = None,
        html: Annotated[bool, Field(description="Treat body as an HTML fragment.")] = False,
        attachments: Annotated[Optional[list[str]], Field(description="Absolute paths of extra files to attach.")] = None,
        subject: Annotated[Optional[str], Field(description="Replace the subject (default keeps 'FW: ...').")] = None,
        importance: Annotated[Optional[str], Field(description="One of: 'low', 'normal', 'high'.")] = None,
        signature: Annotated[SignatureArg, Field(description=SIGNATURE_DESC)] = None,
        save_only: Annotated[bool, Field(description=SAVE_ONLY_DESC)] = False,
    ) -> str:
        """Forward a mail (with its attachments) and an optional note, or save it as a draft."""
        data = await bridge.call(
            mail_client.forward_mail,
            entry_id=entry_id,
            to=to,
            body=body,
            cc=cc,
            bcc=bcc,
            html=html,
            attachments=attachments,
            subject=subject,
            importance=importance,
            signature=signature,
            save_only=save_only,
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_send_draft",
        annotations={
            "title": "Send an existing Outlook draft",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": True,
        },
    )
    @safe_call
    async def outlook_send_draft(
        entry_id: Annotated[str, Field(description="EntryID of the draft (from a save_only result or the Drafts folder).")],
    ) -> str:
        """Send a draft that is already in Drafts, e.g. after the user reviewed it.

        Confirm with the user first. Disabled in draft-only mode (the default; OUTLOOK_MCP_ALLOW_SEND=1 enables sending).
        """
        data = await bridge.call(mail_client.send_draft, entry_id=entry_id)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_list_signatures",
        annotations={
            "title": "List Outlook signatures",
            "readOnlyHint": True,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_list_signatures() -> str:
        """List saved Outlook signatures and the defaults for new mail and replies."""
        data = await bridge.call(mail_client.list_signatures_info)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_move_mail",
        annotations={
            "title": "Move Outlook mail to folder",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_move_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to move.")],
        target_folder: Annotated[str, Field(description="Destination folder.")],
    ) -> str:
        """Move a mail to another folder. Returns the new EntryID."""
        data = await bridge.call(
            mail_client.move_mail, entry_id=entry_id, target_folder=target_folder
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_delete_mail",
        annotations={
            "title": "Delete Outlook mail (move to Deleted Items)",
            "readOnlyHint": False,
            "destructiveHint": True,
            "idempotentHint": False,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_delete_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail to delete.")],
    ) -> str:
        """Delete a mail (Outlook moves it to Deleted Items)."""
        data = await bridge.call(mail_client.delete_mail, entry_id=entry_id)
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_mark_mail",
        annotations={
            "title": "Mark Outlook mail read/unread or flag it",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_mark_mail(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        read: Annotated[Optional[bool], Field(description="True=mark read, False=unread, None=no change.")] = None,
        flagged: Annotated[Optional[bool], Field(description="True=flag for follow-up, False=clear flag.")] = None,
    ) -> str:
        """Toggle read state and/or follow-up flag on a mail."""
        data = await bridge.call(
            mail_client.mark_mail, entry_id=entry_id, read=read, flagged=flagged
        )
        return format_response(data, "json")

    @mcp.tool(
        name="outlook_save_attachments",
        annotations={
            "title": "Save Outlook mail attachments to disk",
            "readOnlyHint": False,
            "destructiveHint": False,
            "idempotentHint": True,
            "openWorldHint": False,
        },
    )
    @safe_call
    async def outlook_save_attachments(
        entry_id: Annotated[str, Field(description="EntryID of the mail.")],
        output_dir: Annotated[str, Field(description="Absolute directory under your user profile.")],
        attachment_index: Annotated[
            Optional[int], Field(ge=1, description="1-indexed attachment. Omit to save all.")
        ] = None,
    ) -> str:
        """Save one or all attachments from a mail to a local directory."""
        data = await bridge.call(
            mail_client.save_attachments,
            entry_id=entry_id,
            output_dir=output_dir,
            attachment_index=attachment_index,
        )
        return format_response(data, "json")
