#!/usr/bin/env python3
"""
Generate time tracking events for November 2025 and store in SQLite database.
"""

import sqlite3
import random
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
from faker import Faker

# Initialize Faker
fake = Faker()

# Database file
DB_FILE = Path(__file__).parent / "events.db"

# Timezone
EST = ZoneInfo("America/New_York")
UTC = ZoneInfo("UTC")

# Projects - Regular (use task/phase codes)
REGULAR_PROJECTS = [
    "Drackett Bay View: 2015.02.2",
    "Finn: 2012.03.1",
    "Wittekind: 2024.04.1",
    "Knight: 2024.xx.x",
    "HHS 857 Cap. Improv. Support (82020-101): 2007.07.5",
    "Lieser: 2017.04.2",
    "Jenkins: 2024.xx.x",
    "Squeri: 2020.12.1",
    "DDB: AP1",
    "CYMI Calvary: 2023.07.1",
    "HHS Fairforest: 2023.08.1",
]

# Office project (max 1hr/week, uses BD task code)
OFFICE_PROJECT = "Office: 2025.01.1"

# Time-off projects (task=NA, phase=NA)
TIME_OFF_PROJECTS = {
    "vacation": "Vacation: 2025.01.2",
    "holiday": "Holiday: 2025.01.3",
    "sick": "Sick: 2025.01.4",
    "personal": "Personal Time: 2025.01.5",
}

# Task codes
TASK_CODES = ["DP", "PM", "3-D", "D-D", "M"]
OFFICE_TASK_CODE = "BD"

# Phase codes
PHASE_CODES = ["PD", "SD", "DD", "CD", "CA", "M"]

# Work descriptions for different task types
WORK_DESCRIPTIONS = {
    "DP": [
        "Design review and feedback session",
        "Reviewed design concepts with team",
        "Principal design guidance",
        "Design direction meeting",
        "Concept development review",
    ],
    "PM": [
        "Project coordination call",
        "Schedule review and updates",
        "Budget tracking and reporting",
        "Client communication",
        "Team resource planning",
        "Project status meeting",
    ],
    "3-D": [
        "3D model development",
        "Revit modeling work",
        "Model coordination",
        "BIM updates",
        "3D visualization work",
    ],
    "D-D": [
        "Drawing set development",
        "Detail drawings",
        "Plan revisions",
        "Section development",
        "Construction document updates",
        "Specification writing",
    ],
    "M": [
        "Client meeting",
        "Contractor coordination meeting",
        "Team meeting",
        "Design review meeting",
        "Progress meeting",
    ],
    "BD": [
        "Business development meeting",
        "Marketing materials review",
        "Proposal preparation",
        "Client outreach",
    ],
}

TIME_OFF_DESCRIPTIONS = {
    "vacation": ["Vacation day", "PTO", "Annual leave", ""],
    "holiday": ["Thanksgiving Day", "Holiday observance", ""],
    "sick": ["Sick day", "Not feeling well", "Medical appointment", ""],
    "personal": ["Personal time", "Personal appointment", "Personal day", ""],
}


def create_database():
    """Create SQLite database and gen_events table."""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Drop table if exists for clean slate
    cursor.execute("DROP TABLE IF EXISTS gen_events")

    # Create table
    cursor.execute("""
        CREATE TABLE gen_events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            event_title TEXT NOT NULL,
            start_timestamp TEXT NOT NULL,
            end_timestamp TEXT NOT NULL,
            wid TEXT,
            task TEXT NOT NULL,
            phase TEXT NOT NULL
        )
    """)

    conn.commit()
    return conn


def get_workdays_november_2025():
    """Get all workdays in November 2025, grouped by week."""
    thanksgiving = datetime(2025, 11, 27)  # Thursday
    workdays_by_week = {}

    for day in range(1, 31):
        date = datetime(2025, 11, day)
        # Skip weekends (5=Saturday, 6=Sunday)
        if date.weekday() >= 5:
            continue
        # Mark Thanksgiving separately
        week_num = date.isocalendar()[1]
        if week_num not in workdays_by_week:
            workdays_by_week[week_num] = []
        workdays_by_week[week_num].append((date, date == thanksgiving))

    return workdays_by_week


def est_to_utc(est_datetime):
    """Convert EST datetime to UTC."""
    est_aware = est_datetime.replace(tzinfo=EST)
    return est_aware.astimezone(UTC)


def generate_event(date, start_hour, duration_hours, project, task, phase, wid):
    """Generate a single event dict."""
    # Handle fractional hours (e.g., 8.5 = 8:30)
    hour = int(start_hour)
    minute = int((start_hour - hour) * 60)
    start_est = datetime(date.year, date.month, date.day, hour, minute, 0)
    end_est = start_est + timedelta(hours=duration_hours)

    start_utc = est_to_utc(start_est)
    end_utc = est_to_utc(end_est)

    return {
        "event_title": project,
        "start_timestamp": start_utc.isoformat(),
        "end_timestamp": end_utc.isoformat(),
        "wid": wid,
        "task": task,
        "phase": phase,
    }


def get_work_description(task_code, project_type=None):
    """Get a realistic work description."""
    if project_type and project_type in TIME_OFF_DESCRIPTIONS:
        descriptions = TIME_OFF_DESCRIPTIONS[project_type]
        return random.choice(descriptions)
    if task_code in WORK_DESCRIPTIONS:
        return random.choice(WORK_DESCRIPTIONS[task_code])
    return fake.sentence(nb_words=6)


def generate_events():
    """Generate all events for November 2025."""
    events = []
    workdays_by_week = get_workdays_november_2025()

    # Pre-schedule guaranteed time-off events to ensure variety
    # Pick specific days for each time-off type
    all_workdays = []
    for week_num, days in workdays_by_week.items():
        for date, is_thanksgiving in days:
            if not is_thanksgiving:
                all_workdays.append(date)

    # Guarantee at least one of each time-off type (excluding holiday which is Thanksgiving)
    guaranteed_time_off = {}
    available_days = all_workdays.copy()
    random.shuffle(available_days)

    # Schedule 1 vacation day, 1 sick day, 1 personal time
    for time_off_type in ["vacation", "sick", "personal"]:
        if available_days:
            day = available_days.pop()
            guaranteed_time_off[day] = time_off_type

    for week_num, days in workdays_by_week.items():
        weekly_hours = 0
        target_hours = random.randint(40, 50)
        office_hours_this_week = 0

        for date, is_thanksgiving in days:
            if is_thanksgiving:
                # Full day holiday for Thanksgiving (8 hours)
                wid = "Thanksgiving Day"
                # Occasionally leave WID empty for time-off
                if random.random() < 0.3:
                    wid = ""
                events.append(
                    generate_event(
                        date, 8, 8, TIME_OFF_PROJECTS["holiday"], "NA", "NA", wid
                    )
                )
                weekly_hours += 8
                continue

            # Generate work events for this day
            daily_hours = 0
            current_hour = 8  # Start at 8 AM EST

            # Check if this is a guaranteed time-off day
            if date in guaranteed_time_off:
                time_off_type = guaranteed_time_off[date]
                hours = random.choice([2, 4, 8])
                wid = get_work_description(None, time_off_type)
                # Occasionally leave WID empty for time-off
                if random.random() < 0.3:
                    wid = ""
                events.append(
                    generate_event(
                        date,
                        current_hour,
                        hours,
                        TIME_OFF_PROJECTS[time_off_type],
                        "NA",
                        "NA",
                        wid,
                    )
                )
                daily_hours += hours
                current_hour += hours
                weekly_hours += hours
                if hours >= 8:
                    continue  # Full day off

            # Maybe add additional random time-off (rare)
            elif random.random() < 0.05:  # 5% chance
                time_off_type = random.choice(["vacation", "sick", "personal"])
                hours = random.choice([2, 4, 8])
                wid = get_work_description(None, time_off_type)
                # Occasionally leave WID empty for time-off
                if random.random() < 0.3:
                    wid = ""
                events.append(
                    generate_event(
                        date,
                        current_hour,
                        hours,
                        TIME_OFF_PROJECTS[time_off_type],
                        "NA",
                        "NA",
                        wid,
                    )
                )
                daily_hours += hours
                current_hour += hours
                weekly_hours += hours
                if hours >= 8:
                    continue  # Full day off

            # Generate work events until day is full or target reached
            while current_hour < 17 and daily_hours < 9:
                remaining_day = 17 - current_hour
                remaining_week = target_hours - weekly_hours

                if remaining_week <= 0:
                    break

                # Decide event type
                event_type = random.choices(
                    ["regular", "meeting", "office"],
                    weights=[0.7, 0.25, 0.05],
                    k=1,
                )[0]

                if event_type == "office" and office_hours_this_week >= 1:
                    event_type = "regular"  # Already hit office cap

                # Duration: 0.5 to 4 hours, in 0.5 increments
                max_duration = min(remaining_day, remaining_week, 4)
                if max_duration < 0.5:
                    break
                duration = random.choice(
                    [d / 2 for d in range(1, int(max_duration * 2) + 1)]
                )

                if event_type == "meeting":
                    project = random.choice(REGULAR_PROJECTS)
                    task = "M"
                    phase = "M"
                    wid = get_work_description("M")
                elif event_type == "office":
                    # Cap office at remaining hours up to 1hr/week max
                    remaining_office = 1 - office_hours_this_week
                    duration = min(duration, remaining_office)
                    project = OFFICE_PROJECT
                    task = OFFICE_TASK_CODE
                    phase = random.choice(["PD", "SD"])  # Office typically early phases
                    wid = get_work_description("BD")
                    office_hours_this_week += duration

                else:  # regular project work
                    project = random.choice(REGULAR_PROJECTS)
                    task = random.choice(TASK_CODES)
                    # Match phase to task somewhat logically
                    if task == "M":
                        phase = "M"
                    else:
                        phase = random.choice(PHASE_CODES)
                    wid = get_work_description(task)

                events.append(
                    generate_event(date, current_hour, duration, project, task, phase, wid)
                )

                current_hour += duration
                daily_hours += duration
                weekly_hours += duration

    return events


def insert_events(conn, events):
    """Insert all events into the database."""
    cursor = conn.cursor()
    cursor.executemany(
        """
        INSERT INTO gen_events (event_title, start_timestamp, end_timestamp, wid, task, phase)
        VALUES (:event_title, :start_timestamp, :end_timestamp, :wid, :task, :phase)
    """,
        events,
    )
    conn.commit()


def print_summary(conn):
    """Print summary statistics."""
    cursor = conn.cursor()

    # Total events
    cursor.execute("SELECT COUNT(*) FROM gen_events")
    total = cursor.fetchone()[0]
    print(f"\nTotal events generated: {total}")

    # Events by project
    print("\nEvents by project:")
    cursor.execute("""
        SELECT event_title, COUNT(*) as count
        FROM gen_events
        GROUP BY event_title
        ORDER BY count DESC
    """)
    for row in cursor.fetchall():
        print(f"  {row[0]}: {row[1]}")

    # Hours by week (approximate based on timestamps)
    print("\nApproximate hours by week:")
    cursor.execute("""
        SELECT
            strftime('%W', start_timestamp) as week,
            ROUND(SUM(
                (julianday(end_timestamp) - julianday(start_timestamp)) * 24
            ), 1) as hours
        FROM gen_events
        GROUP BY week
        ORDER BY week
    """)
    for row in cursor.fetchall():
        print(f"  Week {row[0]}: {row[1]} hours")

    # Task distribution
    print("\nEvents by task code:")
    cursor.execute("""
        SELECT task, COUNT(*) as count
        FROM gen_events
        GROUP BY task
        ORDER BY count DESC
    """)
    for row in cursor.fetchall():
        print(f"  {row[0]}: {row[1]}")

    # Phase distribution
    print("\nEvents by phase code:")
    cursor.execute("""
        SELECT phase, COUNT(*) as count
        FROM gen_events
        GROUP BY phase
        ORDER BY count DESC
    """)
    for row in cursor.fetchall():
        print(f"  {row[0]}: {row[1]}")


def main():
    print("Creating database and generating events...")

    # Create database
    conn = create_database()

    # Generate events
    events = generate_events()

    # Insert events
    insert_events(conn, events)

    # Print summary
    print_summary(conn)

    conn.close()
    print(f"\nDatabase saved to: {DB_FILE}")


if __name__ == "__main__":
    main()
