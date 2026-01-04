import asyncio
import os
import sqlite3
from pathlib import Path

from azure.identity import ClientSecretCredential
from dotenv import load_dotenv
from msgraph import GraphServiceClient
from msgraph.generated.models.event import Event
from msgraph.generated.models.date_time_time_zone import DateTimeTimeZone
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.body_type import BodyType

load_dotenv()

MICROSOFT_GRAPH_TENANT_ID = os.environ["MICROSOFT_GRAPH_TENANT_ID"]
MICROSOFT_GRAPH_APP_ID = os.environ["MICROSOFT_GRAPH_APP_ID"]
MICROSOFT_GRAPH_CLIENT_SECRET = os.environ["MICROSOFT_GRAPH_CLIENT_SECRET"]

# Target user and calendar
TARGET_USER = "charles@landslidelogic.com"
TARGET_CALENDAR = "CES TIME CARD"

# Database file
DB_FILE = Path(__file__).parent / "events.db"

# Create credential using azure-identity
credential = ClientSecretCredential(
    tenant_id=MICROSOFT_GRAPH_TENANT_ID,
    client_id=MICROSOFT_GRAPH_APP_ID,
    client_secret=MICROSOFT_GRAPH_CLIENT_SECRET,
)

# Graph client
graph = GraphServiceClient(credentials=credential)


def read_events_from_db():
    """Read all generated events from SQLite database."""
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT event_title, start_timestamp, end_timestamp, wid, task, phase
        FROM gen_events
        ORDER BY start_timestamp
    """)

    events = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return events


def create_calendar_event(db_event):
    """Convert a database event to MS Graph Event object."""
    # Build description with WID, task, and phase (with blank lines for readability)
    body_parts = []
    if db_event["wid"]:
        body_parts.append(f"WID: {db_event['wid']}")
    body_parts.append(f"Task: {db_event['task']}")
    body_parts.append(f"Phase: {db_event['phase']}")
    body_content = "\n\n".join(body_parts)

    event = Event(
        subject=db_event["event_title"],
        start=DateTimeTimeZone(
            date_time=db_event["start_timestamp"].replace("+00:00", ""),
            time_zone="UTC",
        ),
        end=DateTimeTimeZone(
            date_time=db_event["end_timestamp"].replace("+00:00", ""),
            time_zone="UTC",
        ),
        body=ItemBody(
            content_type=BodyType.Text,
            content=body_content,
        ),
    )
    return event


async def find_calendar(user_id, calendar_name):
    """Find a specific calendar by name for a user."""
    calendars_response = await graph.users.by_user_id(user_id).calendars.get()
    calendars = calendars_response.value

    for calendar in calendars:
        if calendar.name == calendar_name:
            return calendar

    return None


async def delete_all_calendar_events(user_id, calendar_id):
    """Delete all events from a calendar (handles pagination)."""
    deleted_count = 0
    error_count = 0

    # Keep deleting until no events remain (API paginates results)
    while True:
        # Get events from the calendar
        events_response = await graph.users.by_user_id(user_id).calendars.by_calendar_id(
            calendar_id
        ).events.get()

        events = events_response.value if events_response.value else []

        if not events:
            break

        print(f"  Deleting batch of {len(events)} events...")

        for event in events:
            try:
                await graph.users.by_user_id(user_id).calendars.by_calendar_id(
                    calendar_id
                ).events.by_event_id(event.id).delete()
                deleted_count += 1
                print(f"  [{deleted_count}] Deleted: {event.subject}")
            except Exception as e:
                error_count += 1
                print(f"  ERROR deleting {event.subject}: {e}")

    return deleted_count, error_count


async def main():
    print(f"Reading events from {DB_FILE}...")
    db_events = read_events_from_db()
    print(f"Found {len(db_events)} events to add")

    # Find the target user
    print(f"\nLooking up user: {TARGET_USER}")
    user = await graph.users.by_user_id(TARGET_USER).get()
    print(f"Found user: {user.display_name} ({user.user_principal_name})")

    # Find the target calendar
    print(f"\nLooking for calendar: {TARGET_CALENDAR}")
    calendar = await find_calendar(user.id, TARGET_CALENDAR)

    if not calendar:
        print(f"ERROR: Calendar '{TARGET_CALENDAR}' not found!")
        print("\nAvailable calendars:")
        calendars_response = await graph.users.by_user_id(user.id).calendars.get()
        for cal in calendars_response.value:
            print(f"  - {cal.name}")
        return

    print(f"Found calendar: {calendar.name} (ID: {calendar.id})")

    # Delete existing events from the calendar
    print(f"\nDeleting existing events from {TARGET_CALENDAR}...")
    deleted, delete_errors = await delete_all_calendar_events(user.id, calendar.id)
    print(f"Deleted {deleted} events ({delete_errors} errors)")

    # Add events to the calendar
    print(f"\nAdding {len(db_events)} events to calendar...")
    success_count = 0
    error_count = 0

    for i, db_event in enumerate(db_events, 1):
        try:
            ms_event = create_calendar_event(db_event)
            await graph.users.by_user_id(user.id).calendars.by_calendar_id(
                calendar.id
            ).events.post(ms_event)
            success_count += 1
            print(f"  [{i}/{len(db_events)}] Added: {db_event['event_title']} ({db_event['start_timestamp'][:10]})")
        except Exception as e:
            error_count += 1
            print(f"  [{i}/{len(db_events)}] ERROR: {db_event['event_title']} - {e}")

    print(f"\nComplete! Added {success_count} events, {error_count} errors")


asyncio.run(main())
