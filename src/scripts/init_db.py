#!/usr/bin/env python3
"""Create the cci-timesheets SQLite3 database with reports and events tables."""

import sqlite3
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.config import DB_PATH


def create_database():
    """Create the database and tables if they don't exist."""
    # Ensure directory exists
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")
    cursor = conn.cursor()

    # Create reports table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS reports (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            type TEXT NOT NULL CHECK(type IN ('weekly_report', 'monthly_report', 'timesheet_weekly_report', 'timesheet_monthly_report')),
            name TEXT UNIQUE NOT NULL,
            create_date TEXT DEFAULT CURRENT_TIMESTAMP
        )
    """)

    # Create events table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS events (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id INTEGER NOT NULL,
            project_id TEXT,
            employee_id TEXT,
            start_timestamp TEXT,
            end_timestamp TEXT,
            task TEXT,
            phase TEXT,
            wid TEXT,
            error_message TEXT,
            FOREIGN KEY (report_id) REFERENCES reports(id)
        )
    """)

    # Create API request logging table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS api_requests (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            request_id TEXT UNIQUE NOT NULL,
            timestamp TEXT NOT NULL,
            endpoint TEXT NOT NULL,
            method TEXT NOT NULL,
            client_ip TEXT,
            file_size_bytes INTEGER,
            file_name TEXT,
            invoice_date_override TEXT,
            status_code INTEGER NOT NULL,
            error_code TEXT,
            error_message TEXT,
            processing_time_ms INTEGER NOT NULL,
            projects_generated INTEGER,
            total_hours REAL
        )
    """)

    # Create API request details table (for validation errors)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS api_request_details (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            request_id TEXT NOT NULL,
            detail_type TEXT NOT NULL CHECK(detail_type IN ('validation_error', 'project_processed', 'warning')),
            message TEXT NOT NULL,
            FOREIGN KEY (request_id) REFERENCES api_requests(request_id)
        )
    """)

    # Create indexes for API logging
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_api_requests_timestamp ON api_requests(timestamp)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_api_requests_status ON api_requests(status_code)"
    )
    cursor.execute(
        "CREATE INDEX IF NOT EXISTS idx_api_request_details_request ON api_request_details(request_id)"
    )

    conn.commit()
    conn.close()
    print(f"Database created successfully at: {DB_PATH}")


if __name__ == "__main__":
    create_database()
