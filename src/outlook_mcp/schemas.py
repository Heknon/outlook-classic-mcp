"""Shared Pydantic models and enums for tool parameters."""

from __future__ import annotations

from enum import Enum
from typing import Literal, Optional

from pydantic import BaseModel, Field


class ResponseFormat(str, Enum):
    JSON = "json"
    MARKDOWN = "markdown"


class Recurrence(BaseModel):
    """Recurrence pattern for calendar events.

    ``type`` is the cadence; ``interval`` is the spacing (every N days /
    weeks / months / years). Provide either ``occurrences`` or
    ``end_date`` to bound the series — if both are omitted the pattern
    runs indefinitely.
    """

    type: Literal["daily", "weekly", "monthly", "yearly"] = Field(
        description="Cadence: daily, weekly, monthly, or yearly."
    )
    interval: int = Field(
        default=1, ge=1, description="Every N units (days/weeks/months/years)."
    )
    occurrences: Optional[int] = Field(
        default=None, ge=1, description="End after N occurrences."
    )
    end_date: Optional[str] = Field(
        default=None,
        description="ISO-8601 date for series end (alternative to occurrences).",
    )


class RuleConditions(BaseModel):
    """When a rule applies (or, as exceptions, when it doesn't).

    All the conditions you set must hold. Word lists match if ANY of the
    words appears (like Outlook's "with specific words").
    """

    from_people: list[str] = Field(default_factory=list, description="Sender is one of these (addresses or names).")
    sent_to: list[str] = Field(default_factory=list, description="Sent to one of these people or groups.")
    subject_words: list[str] = Field(default_factory=list, description="Subject contains any of these.")
    body_words: list[str] = Field(default_factory=list, description="Body contains any of these.")
    subject_or_body_words: list[str] = Field(default_factory=list, description="Subject or body contains any of these.")
    sender_address_words: list[str] = Field(default_factory=list, description="Sender's address contains any of these, e.g. '@vendor.com'.")
    recipient_address_words: list[str] = Field(default_factory=list, description="A recipient address contains any of these.")
    header_words: list[str] = Field(default_factory=list, description="Message header contains any of these.")
    categories: list[str] = Field(default_factory=list, description="Item has one of these categories.")
    importance: Optional[Literal["low", "normal", "high"]] = Field(default=None, description="Marked with this importance.")
    to_me: bool = Field(default=False, description="My name is in To.")
    only_to_me: bool = Field(default=False, description="Sent only to me.")
    cc_me: bool = Field(default=False, description="My name is in CC.")
    to_or_cc_me: bool = Field(default=False, description="My name is in To or CC.")
    has_attachment: bool = Field(default=False, description="Has an attachment.")


class RuleActions(BaseModel):
    """What a rule does. Forward/redirect actions are restricted (see tool description)."""

    move_to_folder: Optional[str] = Field(default=None, description="Move to this folder (e.g. 'Inbox/Vendors').")
    copy_to_folder: Optional[str] = Field(default=None, description="Copy to this folder.")
    categories: list[str] = Field(default_factory=list, description="Assign these categories (created if missing).")
    category_color: Optional[str] = Field(default=None, description="Colour for categories this rule creates.")
    clear_categories: bool = Field(default=False, description="Remove all categories.")
    mark_as_task: Optional[Literal["today", "tomorrow", "this_week", "next_week", "no_date"]] = Field(
        default=None, description="Flag for follow-up."
    )
    delete: bool = Field(default=False, description="Move to Deleted Items.")
    delete_permanently: bool = Field(default=False, description="Delete permanently (no recovery).")
    stop_processing: bool = Field(default=False, description="Stop processing more rules.")
    desktop_alert: bool = Field(default=False, description="Show a desktop alert.")
    forward_to: list[str] = Field(default_factory=list, description="Forward to these people (restricted).")
    forward_as_attachment_to: list[str] = Field(default_factory=list, description="Forward as an attachment (restricted).")
    redirect_to: list[str] = Field(default_factory=list, description="Redirect to these people (restricted).")
