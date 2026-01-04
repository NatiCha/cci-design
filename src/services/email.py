"""
Email sending functionality for reports.
"""

import traceback
from collections import defaultdict
from datetime import date
from pathlib import Path

from msgraph.generated.models.body_type import BodyType
from msgraph.generated.models.email_address import EmailAddress
from msgraph.generated.models.file_attachment import FileAttachment
from msgraph.generated.models.item_body import ItemBody
from msgraph.generated.models.message import Message
from msgraph.generated.models.recipient import Recipient
from msgraph.generated.users.item.send_mail.send_mail_post_request_body import (
    SendMailPostRequestBody,
)

from core.config import ERROR_EMAIL, FROM_EMAIL, TO_EMAIL
from core.graph_client import get_graph_client
from services.reports import format_date_for_subject, format_date_short


def format_conflicts_for_email(events: list[dict], as_of_date: date, report_type: str) -> str:
    """Format conflicts into hierarchical email body."""
    date_str = format_date_for_subject(as_of_date, report_type)
    report_title = "Weekly" if report_type == "weekly_report" else "Monthly"

    lines = [f"Timesheet {report_title} Report - {date_str}", ""]

    # Group conflicts by employee, then by project
    conflicts_by_employee: dict[str, dict[str, list[dict]]] = defaultdict(
        lambda: defaultdict(list)
    )
    employees_with_data = set()

    for event in events:
        employees_with_data.add(event["employee_id"])
        if event["error_message"]:
            conflicts_by_employee[event["employee_id"]][event["project_id"]].append(event)

    if conflicts_by_employee:
        lines.append("Conflicts Found:")
        lines.append("")

        for employee in sorted(conflicts_by_employee.keys()):
            lines.append(f"{employee}:")
            projects = conflicts_by_employee[employee]

            for project_id in sorted(projects.keys()):
                lines.append(f"  {project_id}")
                for event in projects[project_id]:
                    date_str = ""
                    if event["event_date"]:
                        date_str = format_date_short(event["event_date"])
                    lines.append(f"    - {date_str}: {event['error_message']}")
            lines.append("")

    # List employees without conflicts
    employees_without_conflicts = employees_with_data - set(conflicts_by_employee.keys())
    if employees_without_conflicts:
        lines.append(f"No conflicts found for: {', '.join(sorted(employees_without_conflicts))}")

    if not conflicts_by_employee and not employees_with_data:
        lines.append("No events found for this period.")

    return "\n".join(lines)


async def send_report_email(
    report_name: str, file_path: Path, events: list[dict], as_of_date: date, report_type: str
):
    """Send report email with attachment."""
    graph = get_graph_client()
    date_str = format_date_for_subject(as_of_date, report_type)
    report_title = "Weekly" if report_type == "weekly_report" else "Monthly"
    subject = f"Timesheet {report_title} {date_str}"

    body_text = format_conflicts_for_email(events, as_of_date, report_type)

    # Read attachment directly as bytes
    with open(file_path, "rb") as f:
        attachment_bytes = f.read()

    attachment = FileAttachment(
        odata_type="#microsoft.graph.fileAttachment",
        name=file_path.name,
        content_type="application/vnd.apple.numbers",
        content_bytes=attachment_bytes,
    )

    message = Message(
        subject=subject,
        body=ItemBody(content_type=BodyType.Text, content=body_text),
        to_recipients=[Recipient(email_address=EmailAddress(address=TO_EMAIL))],
        attachments=[attachment],
    )

    request_body = SendMailPostRequestBody(message=message, save_to_sent_items=True)

    await graph.users.by_user_id(FROM_EMAIL).send_mail.post(request_body)
    print(f"Sent report email to {TO_EMAIL}")


async def send_error_email(error: Exception):
    """Send error notification email."""
    graph = get_graph_client()
    subject = "Timesheet Report - Script Error"
    body_text = f"An error occurred while generating the timesheet report:\n\n{traceback.format_exc()}"

    message = Message(
        subject=subject,
        body=ItemBody(content_type=BodyType.Text, content=body_text),
        to_recipients=[Recipient(email_address=EmailAddress(address=ERROR_EMAIL))],
    )

    request_body = SendMailPostRequestBody(message=message, save_to_sent_items=True)

    try:
        await graph.users.by_user_id(FROM_EMAIL).send_mail.post(request_body)
        print(f"Sent error email to {ERROR_EMAIL}")
    except Exception as e:
        print(f"Failed to send error email: {e}")
