"""
SQLite database operations for timesheet reports.
"""

import sqlite3
from datetime import date

from core.config import DB_PATH


def get_connection() -> sqlite3.Connection:
    """Get a database connection."""
    return sqlite3.connect(DB_PATH)


def generate_report_name(report_type: str, as_of_date: date, conn: sqlite3.Connection) -> str:
    """
    Generate unique report name with auto-incremented suffix.

    Example: timesheet_weekly_report_2025_11_07_a, timesheet_monthly_report_2025_11_a
    """
    # Monthly reports use YYYY_MM format, weekly reports use YYYY_MM_DD format
    if "monthly" in report_type:
        date_str = as_of_date.strftime("%Y_%m")
    else:
        date_str = as_of_date.strftime("%Y_%m_%d")
    base_pattern = f"{report_type}_{date_str}_"

    cursor = conn.cursor()
    cursor.execute(
        "SELECT name FROM reports WHERE name LIKE ? ORDER BY name DESC",
        (f"{base_pattern}%",),
    )
    existing = cursor.fetchall()

    if not existing:
        return f"{base_pattern}a"

    # Find the highest suffix
    highest_suffix = "a"
    for (name,) in existing:
        suffix = name.replace(base_pattern, "")
        if suffix and suffix > highest_suffix:
            highest_suffix = suffix

    # Increment suffix
    next_suffix = chr(ord(highest_suffix) + 1)
    return f"{base_pattern}{next_suffix}"


def create_report_record(
    conn: sqlite3.Connection, report_type: str, report_name: str
) -> int:
    """Create report record and return report_id."""
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO reports (type, name) VALUES (?, ?)",
        (report_type, report_name),
    )
    conn.commit()
    return cursor.lastrowid


def insert_events(conn: sqlite3.Connection, report_id: int, events: list[dict]):
    """Insert all events linked to report_id."""
    cursor = conn.cursor()
    for event in events:
        cursor.execute(
            """
            INSERT INTO events (
                report_id, project_id, employee_id, start_timestamp,
                end_timestamp, task, phase, wid, error_message
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                report_id,
                event["project_id"],
                event["employee_id"],
                event["start_timestamp"],
                event["end_timestamp"],
                event["task"],
                event["phase"],
                event["wid"],
                event["error_message"],
            ),
        )
    conn.commit()
