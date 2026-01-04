#!/usr/bin/env python3
"""
Create monthly timesheet report from MS365 calendars.

Generates Excel report with three sheets:
- View: Standard timesheet detail (read-only reference)
- Edit: Timesheet detail with Hours Adjusted and Total Adjusted Hours columns
- Billable Goals: Dynamic employee columns with SUMIF formulas for billing metrics

Usage:
    uv run python src/scripts/create_monthly_report.py --month 2025-11
"""

import argparse
import asyncio
import calendar
import sqlite3
import sys
import traceback
from datetime import date, datetime
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.config import DB_PATH, OUTPUT_DIR
from core.database import create_report_record, generate_report_name, insert_events
from core.validation import validate_events
from services.calendar import discover_time_card_calendars, fetch_calendar_events
from services.email import send_error_email, send_report_email
from services.reports import create_monthly_excel_report


# =============================================================================
# DATE UTILITIES
# =============================================================================


def get_monthly_date_range(month_str: str | None) -> tuple[date, date]:
    """
    Calculate date range for monthly report.

    Args:
        month_str: Optional month string (YYYY-MM). Uses previous month if None.

    Returns:
        Tuple of (first_of_month, last_of_month)
    """
    if month_str:
        # Parse YYYY-MM format
        year, month = map(int, month_str.split("-"))
        target_date = date(year, month, 1)
    else:
        # Default to previous month
        today = date.today()
        if today.month == 1:
            target_date = date(today.year - 1, 12, 1)
        else:
            target_date = date(today.year, today.month - 1, 1)

    first_of_month = target_date.replace(day=1)
    _, last_day = calendar.monthrange(target_date.year, target_date.month)
    last_of_month = target_date.replace(day=last_day)

    return first_of_month, last_of_month


# =============================================================================
# MAIN
# =============================================================================


async def main(month_str: str | None = None):
    """Main entry point for monthly report."""
    try:
        # 1. Calculate date range (full month)
        start_date, end_date = get_monthly_date_range(month_str)
        print(f"Generating monthly report for {start_date} to {end_date}")

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
        report_name = generate_report_name("timesheet_monthly_report", end_date, conn)
        report_id = create_report_record(conn, "timesheet_monthly_report", report_name)
        insert_events(conn, report_id, validated_events)
        conn.close()
        print(f"\nCreated report: {report_name} (ID: {report_id})")

        # 6. Generate Excel file (View, Edit, and Billable Goals sheets)
        output_dir = OUTPUT_DIR / "reports" / "monthly"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = output_dir / f"{report_name}.xlsx"
        create_monthly_excel_report(validated_events, output_path)

        # 7. Send email with report and conflicts
        await send_report_email(
            report_name, output_path, validated_events, end_date, "timesheet_monthly_report"
        )

        print("\nDone!")

    except Exception as e:
        print(f"\nError: {e}")
        traceback.print_exc()
        await send_error_email(e)
        raise


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate monthly timesheet report")
    parser.add_argument(
        "--month",
        help="Target month (YYYY-MM). Defaults to previous month.",
    )
    args = parser.parse_args()

    asyncio.run(main(args.month))
