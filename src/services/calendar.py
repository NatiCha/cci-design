"""
Calendar discovery and event fetching from MS Graph.
"""

import re
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from core.config import CALENDAR_PATTERN
from core.graph_client import get_graph_client


async def discover_time_card_calendars() -> list[dict]:
    """
    Scan all org users for calendars matching 'XXX TIME CARD' pattern.

    Returns:
        List of dicts with user_id, calendar_id, initials, user_email
    """
    graph = get_graph_client()
    calendars_found = []

    # Get all users in the organization
    users_response = await graph.users.get()
    users = users_response.value if users_response.value else []

    print(f"Scanning {len(users)} users for TIME CARD calendars...")

    for user in users:
        try:
            calendars_response = await graph.users.by_user_id(user.id).calendars.get()
            calendars = calendars_response.value if calendars_response.value else []

            for calendar in calendars:
                if calendar.name and CALENDAR_PATTERN.lower() in calendar.name.lower():
                    # Extract initials (e.g., "CES" from "CES TIME CARD" or "CES Time Card")
                    initials = re.sub(CALENDAR_PATTERN, "", calendar.name, flags=re.IGNORECASE).strip().upper()
                    calendars_found.append(
                        {
                            "user_id": user.id,
                            "user_email": user.user_principal_name,
                            "calendar_id": calendar.id,
                            "calendar_name": calendar.name,
                            "initials": initials,
                        }
                    )
                    print(f"  Found: {calendar.name} ({user.user_principal_name})")
        except Exception as e:
            # Silently skip users without mailboxes (service accounts, admin accounts, etc.)
            error_code = getattr(getattr(e, "error", None), "code", None)
            if error_code == "MailboxNotEnabledForRESTAPI":
                continue
            print(f"  Error scanning {user.user_principal_name}: {e}")

    return calendars_found


async def fetch_calendar_events(
    user_id: str, calendar_id: str, initials: str, start_date: date, end_date: date
) -> list[dict]:
    """
    Fetch all events from a calendar within date range.

    Handles pagination and parses event body for WID, Task, Phase.
    """
    graph = get_graph_client()
    events = []

    # Convert dates to datetime with timezone for API
    start_dt = datetime.combine(start_date, datetime.min.time()).replace(
        tzinfo=ZoneInfo("UTC")
    )
    # End date should include the full day
    end_dt = datetime.combine(end_date + timedelta(days=1), datetime.min.time()).replace(
        tzinfo=ZoneInfo("UTC")
    )

    # Build filter for date range
    start_str = start_dt.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str = end_dt.strftime("%Y-%m-%dT%H:%M:%SZ")

    try:
        # Fetch events with filter
        from msgraph.generated.users.item.calendars.item.events.events_request_builder import (
            EventsRequestBuilder,
        )

        query_params = EventsRequestBuilder.EventsRequestBuilderGetQueryParameters(
            filter=f"start/dateTime ge '{start_str}' and start/dateTime lt '{end_str}'",
            orderby=["start/dateTime"],
            top=100,
        )
        config = EventsRequestBuilder.EventsRequestBuilderGetRequestConfiguration(
            query_parameters=query_params
        )

        events_response = await graph.users.by_user_id(user_id).calendars.by_calendar_id(
            calendar_id
        ).events.get(request_configuration=config)

        raw_events = events_response.value if events_response.value else []

        for event in raw_events:
            parsed = parse_event(event, initials)
            events.append(parsed)

    except Exception as e:
        print(f"  Error fetching events: {e}")

    return events


def parse_event(event, initials: str) -> dict:
    """Parse MS Graph event into our format."""
    # Parse body for WID, Task, Phase
    wid = ""
    task = ""
    phase = ""

    if event.body and event.body.content:
        body_text = event.body.content.strip()
        # Handle both plain text and HTML
        if "<" in body_text:
            # Simple HTML stripping
            body_text = re.sub(r"<[^>]+>", "\n", body_text)

        for line in body_text.split("\n"):
            line = line.strip()
            if line.upper().startswith("WID:"):
                wid = line[4:].strip()
            elif line.upper().startswith("TASK:"):
                task = line[5:].strip().upper()
            elif line.upper().startswith("PHASE:"):
                phase = line[6:].strip().upper()

    # Parse timestamps
    start_ts = None
    end_ts = None
    if event.start and event.start.date_time:
        start_ts = event.start.date_time
    if event.end and event.end.date_time:
        end_ts = event.end.date_time

    # Calculate hours
    hours = 0.0
    if start_ts and end_ts:
        try:
            start_dt = datetime.fromisoformat(start_ts.replace("Z", "+00:00"))
            end_dt = datetime.fromisoformat(end_ts.replace("Z", "+00:00"))
            hours = round((end_dt - start_dt).total_seconds() / 3600, 2)
        except Exception:
            pass

    # Extract date for display
    event_date = None
    if start_ts:
        try:
            start_dt = datetime.fromisoformat(start_ts.replace("Z", "+00:00"))
            event_date = start_dt.date()
        except Exception:
            pass

    return {
        "project_id": event.subject or "",
        "employee_id": initials,
        "start_timestamp": start_ts,
        "end_timestamp": end_ts,
        "event_date": event_date,
        "hours": hours,
        "task": task,
        "phase": phase,
        "wid": wid,
        "error_message": None,
    }
