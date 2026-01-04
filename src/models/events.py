"""
Data models for events and reports.

Currently using TypedDict for type hints on event dictionaries.
Can be expanded to dataclasses or Pydantic models as needed.
"""

from datetime import date
from typing import TypedDict


class CalendarInfo(TypedDict):
    """Calendar discovery result."""
    user_id: str
    user_email: str
    calendar_id: str
    calendar_name: str
    initials: str


class Event(TypedDict):
    """Parsed calendar event."""
    project_id: str
    employee_id: str
    start_timestamp: str | None
    end_timestamp: str | None
    event_date: date | None
    hours: float
    task: str
    phase: str
    wid: str
    error_message: str | None
