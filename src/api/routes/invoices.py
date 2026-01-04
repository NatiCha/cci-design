"""Invoice generation endpoint."""

import asyncio
import tempfile
import time
from datetime import date, datetime
from pathlib import Path
from typing import Annotated

from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    HTTPException,
    Request,
    UploadFile,
    status,
)
from fastapi.responses import Response

from api.dependencies import verify_api_key
from api.logging import RequestLog, log_request
from api.models.responses import ErrorCodes
from core.config import MAX_UPLOAD_SIZE_BYTES
from services.invoices import generate_invoices_to_bytes

router = APIRouter(prefix="/v1")


def get_client_ip(request: Request) -> str:
    """Extract client IP from request, handling proxies."""
    forwarded = request.headers.get("X-Forwarded-For")
    if forwarded:
        return forwarded.split(",")[0].strip()
    return request.client.host if request.client else "unknown"


def parse_invoice_date(date_str: str | None) -> date | None:
    """Parse invoice date string to date object."""
    if not date_str:
        return None
    try:
        return datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail={
                "error": "Invalid invoice_date format",
                "code": ErrorCodes.INVALID_REQUEST,
                "details": ["Expected format: YYYY-MM-DD"],
            },
        )


def _process_in_thread(
    file_content: bytes,
    invoice_date: date | None,
) -> tuple[bytes, str, int, float]:
    """
    Process invoice generation in thread pool.

    Writes content to temp file (required by numbers-parser),
    calls service, returns bytes.
    """
    # Use delete=False and explicit cleanup to prevent race conditions
    tmp = tempfile.NamedTemporaryFile(suffix=".numbers", delete=False)
    tmp_path = Path(tmp.name)
    try:
        tmp.write(file_content)
        tmp.flush()
        tmp.close()  # Close before reading to ensure data is flushed

        return generate_invoices_to_bytes(tmp_path, invoice_date)
    finally:
        tmp_path.unlink(missing_ok=True)


@router.post("/invoices/generate")
async def generate_invoices_endpoint(
    request: Request,
    file: Annotated[
        UploadFile, File(description="Apple Numbers monthly report file")
    ],
    invoice_date: Annotated[
        str | None, Form(description="Override invoice date (YYYY-MM-DD)")
    ] = None,
    _api_key: str = Depends(verify_api_key),
):
    """
    Generate invoices from a monthly timesheet report.

    Accepts an Apple Numbers file upload and returns an Excel workbook.
    """
    start_time = time.time()

    # Initialize request log
    request_log = RequestLog(
        endpoint="/v1/invoices/generate",
        method="POST",
        client_ip=get_client_ip(request),
        file_name=file.filename,
        invoice_date_override=invoice_date,
    )

    try:
        # Validate file presence
        if not file or not file.filename:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail={
                    "error": "No file provided",
                    "code": ErrorCodes.INVALID_REQUEST,
                    "details": [],
                },
            )

        # Validate file extension
        if not file.filename.endswith(".numbers"):
            raise HTTPException(
                status_code=status.HTTP_415_UNSUPPORTED_MEDIA_TYPE,
                detail={
                    "error": "File is not an Apple Numbers document",
                    "code": ErrorCodes.UNSUPPORTED_MEDIA_TYPE,
                    "details": [f"Received: {file.filename}"],
                },
            )

        # Read file content
        file_content = await file.read()
        request_log.file_size_bytes = len(file_content)

        # Validate file size
        if len(file_content) > MAX_UPLOAD_SIZE_BYTES:
            max_mb = MAX_UPLOAD_SIZE_BYTES // (1024 * 1024)
            raise HTTPException(
                status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                detail={
                    "error": f"File exceeds maximum size of {max_mb} MB",
                    "code": ErrorCodes.FILE_TOO_LARGE,
                    "details": [f"File size: {len(file_content) / (1024*1024):.1f} MB"],
                },
            )

        # Parse invoice date if provided
        parsed_date = parse_invoice_date(invoice_date)

        # Use thread pool for sync file I/O and processing
        excel_bytes, output_filename, project_count, total_hours = await asyncio.to_thread(
            _process_in_thread,
            file_content,
            parsed_date,
        )

        # Log success
        request_log.status_code = 200
        request_log.projects_generated = project_count
        request_log.total_hours = total_hours
        request_log.processing_time_ms = int((time.time() - start_time) * 1000)

        # Return Excel file
        return Response(
            content=excel_bytes,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{output_filename}"'},
        )

    except HTTPException as e:
        # Log HTTP errors
        request_log.status_code = e.status_code
        if isinstance(e.detail, dict):
            request_log.error_code = e.detail.get("code")
            request_log.error_message = e.detail.get("error")
            for detail in e.detail.get("details", []):
                request_log.details.append(("validation_error", detail))
        else:
            request_log.error_message = str(e.detail)
        request_log.processing_time_ms = int((time.time() - start_time) * 1000)
        raise

    except ValueError as e:
        # Validation errors from service
        error_msg = str(e)
        request_log.status_code = 422
        request_log.error_code = ErrorCodes.VALIDATION_ERROR
        request_log.error_message = error_msg
        request_log.processing_time_ms = int((time.time() - start_time) * 1000)

        # Parse multi-line error messages
        details = error_msg.split("\n") if "\n" in error_msg else [error_msg]
        for detail in details:
            if detail.strip():
                request_log.details.append(("validation_error", detail.strip()))

        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail={
                "error": "Numbers file validation failed",
                "code": ErrorCodes.VALIDATION_ERROR,
                "details": [d.strip() for d in details if d.strip()],
            },
        )

    except FileNotFoundError as e:
        request_log.status_code = 500
        request_log.error_code = ErrorCodes.INTERNAL_ERROR
        request_log.error_message = str(e)
        request_log.processing_time_ms = int((time.time() - start_time) * 1000)

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail={
                "error": "Server configuration error",
                "code": ErrorCodes.INTERNAL_ERROR,
                "details": [str(e)],
            },
        )

    except Exception as e:
        # Unexpected errors
        request_log.status_code = 500
        request_log.error_code = ErrorCodes.INTERNAL_ERROR
        request_log.error_message = str(e)
        request_log.processing_time_ms = int((time.time() - start_time) * 1000)

        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail={
                "error": "Internal server error",
                "code": ErrorCodes.INTERNAL_ERROR,
                "details": [],
            },
        )

    finally:
        # Always log the request
        try:
            log_request(request_log)
        except Exception:
            # Don't fail the request if logging fails
            pass
