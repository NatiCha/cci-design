"""SQLite request logging for API."""

import sqlite3
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone

from core.config import DB_PATH


@dataclass
class RequestLog:
    """Captured request/response data for logging."""

    request_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    timestamp: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )
    endpoint: str = ""
    method: str = ""
    client_ip: str | None = None
    file_size_bytes: int | None = None
    file_name: str | None = None
    invoice_date_override: str | None = None
    status_code: int = 0
    error_code: str | None = None
    error_message: str | None = None
    processing_time_ms: int = 0
    projects_generated: int | None = None
    total_hours: float | None = None
    details: list[tuple[str, str]] = field(default_factory=list)  # (type, message)


def log_request(log: RequestLog) -> None:
    """Write request log to SQLite database."""
    conn = sqlite3.connect(DB_PATH)
    try:
        cursor = conn.cursor()

        # Insert main request record
        cursor.execute(
            """
            INSERT INTO api_requests (
                request_id, timestamp, endpoint, method, client_ip,
                file_size_bytes, file_name, invoice_date_override,
                status_code, error_code, error_message, processing_time_ms,
                projects_generated, total_hours
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
            (
                log.request_id,
                log.timestamp,
                log.endpoint,
                log.method,
                log.client_ip,
                log.file_size_bytes,
                log.file_name,
                log.invoice_date_override,
                log.status_code,
                log.error_code,
                log.error_message,
                log.processing_time_ms,
                log.projects_generated,
                log.total_hours,
            ),
        )

        # Insert detail records
        for detail_type, message in log.details:
            cursor.execute(
                """
                INSERT INTO api_request_details (request_id, detail_type, message)
                VALUES (?, ?, ?)
            """,
                (log.request_id, detail_type, message),
            )

        conn.commit()
    finally:
        conn.close()
