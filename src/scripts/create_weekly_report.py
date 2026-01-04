#!/usr/bin/env python3
"""
Create weekly timesheet report from MS365 calendars.

Scans all org users for 'XXX TIME CARD' calendars, fetches events,
validates data, generates Numbers report, and sends email.

Usage:
    uv run python src/scripts/create_weekly_report.py --date 2025-11-07
"""

import argparse
import asyncio
import sqlite3
import sys
import traceback
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from numbers_parser import Document

from core.config import DB_PATH, DETAIL_HEADERS, MIN_TABLE_ROWS, OUTPUT_DIR
from core.database import create_report_record, generate_report_name, insert_events
from core.validation import validate_events
from services.calendar import discover_time_card_calendars, fetch_calendar_events
from services.email import send_error_email, send_report_email
from services.reports import format_date_display, write_detail_table


# =============================================================================
# DATE UTILITIES
# =============================================================================


def get_weekly_date_range(as_of_date_str: str | None) -> tuple[date, date]:
    """
    Calculate date range for weekly report.

    Args:
        as_of_date_str: Optional date string (YYYY-MM-DD). Uses today if None.

    Returns:
        Tuple of (first_of_month, as_of_date)
    """
    if as_of_date_str:
        as_of = datetime.strptime(as_of_date_str, "%Y-%m-%d").date()
    else:
        as_of = date.today()

    first_of_month = as_of.replace(day=1)
    return first_of_month, as_of


# =============================================================================
# NUMBERS FILE GENERATION
# =============================================================================


def create_weekly_numbers_report(events: list[dict], output_path: Path):
    """
    Create Numbers report with Summary and Detail sheets.

    Sheet 1 - Timesheet Summary: Pivot-style table with Project IDs as rows,
              Employee initials as columns, and Total column.
    Sheet 2 - Timesheet Detail: Raw event data with all columns.
    """
    # Get unique employees and projects, calculate hours by project/employee
    employees = sorted(set(e["employee_id"].upper() for e in events))
    projects = sorted(set(e["project_id"] for e in events))

    # Build hours matrix: project -> employee -> total hours
    hours_matrix: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for event in events:
        project = event["project_id"]
        employee = event["employee_id"].upper()
        hours_matrix[project][employee] += event["hours"]

    # ==========================================================================
    # Sheet 1: Timesheet Summary (pivot-style)
    # ==========================================================================
    summary_headers = ["Project ID"] + employees + ["Total"]
    summary_rows = len(projects) + 2  # header + data + total row

    # Create document with Summary sheet/table names from the start
    doc = Document(
        sheet_name="Timesheet Summary",
        table_name="Timesheet Summary",
        num_rows=summary_rows,
        num_cols=len(summary_headers),
        num_header_rows=1,
        num_header_cols=1,
    )

    summary_table = doc.sheets["Timesheet Summary"].tables["Timesheet Summary"]

    # Write headers
    for col_idx, header in enumerate(summary_headers):
        summary_table.write(0, col_idx, header)

    # Write project rows
    for row_idx, project in enumerate(projects, start=1):
        summary_table.write(row_idx, 0, project)
        row_total = 0.0
        for col_idx, employee in enumerate(employees, start=1):
            hours = hours_matrix[project].get(employee, 0.0)
            if hours > 0:
                summary_table.write(row_idx, col_idx, hours)
            row_total += hours
        # Total column
        summary_table.write(row_idx, len(employees) + 1, row_total)

    # Write totals row
    total_row_idx = len(projects) + 1
    summary_table.write(total_row_idx, 0, "Total")
    grand_total = 0.0
    for col_idx, employee in enumerate(employees, start=1):
        emp_total = sum(hours_matrix[p].get(employee, 0.0) for p in projects)
        summary_table.write(total_row_idx, col_idx, emp_total)
        grand_total += emp_total
    summary_table.write(total_row_idx, len(employees) + 1, grand_total)

    # ==========================================================================
    # Sheet 2: Timesheet Detail
    # ==========================================================================
    detail_rows = max(len(events) + 1, MIN_TABLE_ROWS)

    # Add sheet with properly named table
    doc.add_sheet(
        "Timesheet Detail",
        table_name="Timesheet Detail",
        num_rows=detail_rows,
        num_cols=len(DETAIL_HEADERS),
    )
    detail_table = doc.sheets["Timesheet Detail"].tables["Timesheet Detail"]

    # Write headers and data using shared helper
    write_detail_table(detail_table, events)

    # Save to output path
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Saved Numbers report to: {output_path}")


# =============================================================================
# MAIN
# =============================================================================


async def main(as_of_date_str: str | None = None):
    """Main entry point."""
    try:
        # 1. Calculate date range
        start_date, end_date = get_weekly_date_range(as_of_date_str)
        print(f"Generating report for {start_date} to {end_date}")

        # 2. Discover users with TIME CARD calendars
        calendars = await discover_time_card_calendars()

        if not calendars:
            print("No TIME CARD calendars found!")
            return

        # 3. Fetch events from all calendars
        print(f"\nFetching events from {len(calendars)} calendar(s)...")
        all_events = []
        for cal in calendars:
            print(f"  Fetching from {cal['calendar_name']}...")
            events = await fetch_calendar_events(
                cal["user_id"],
                cal["calendar_id"],
                cal["initials"],
                start_date,
                end_date,
            )
            print(f"    Found {len(events)} events")
            all_events.extend(events)

        print(f"\nTotal events: {len(all_events)}")

        # 4. Validate events (conflict resolution)
        validated_events = validate_events(all_events)
        conflicts = [e for e in validated_events if e.get("error_message")]
        print(f"Events with conflicts: {len(conflicts)}")

        # 5. Create database records
        conn = sqlite3.connect(DB_PATH)
        report_name = generate_report_name("timesheet_weekly_report", end_date, conn)
        report_id = create_report_record(conn, "timesheet_weekly_report", report_name)
        insert_events(conn, report_id, validated_events)
        conn.close()
        print(f"\nCreated report: {report_name} (ID: {report_id})")

        # 6. Generate Numbers file
        output_dir = OUTPUT_DIR / "reports" / "weekly"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{report_name}.numbers"
        create_weekly_numbers_report(validated_events, output_path)

        # 7. Send email with report and conflicts
        await send_report_email(
            report_name, output_path, validated_events, end_date, "timesheet_weekly_report"
        )

        print("\nDone!")

    except Exception as e:
        print(f"\nError: {e}")
        traceback.print_exc()
        await send_error_email(e)
        raise


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate weekly timesheet report")
    parser.add_argument(
        "--date",
        help="As-of date (YYYY-MM-DD). Pulls from 1st of month to this date. Defaults to today.",
    )
    args = parser.parse_args()

    asyncio.run(main(args.date))
