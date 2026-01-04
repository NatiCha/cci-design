"""Health check endpoint."""

from datetime import datetime, timezone

from fastapi import APIRouter
from fastapi.responses import JSONResponse

from api.models.responses import HealthResponse
from core.config import API_VERSION
from services.invoices import TEMPLATE_PATH

router = APIRouter()


@router.get("/health", response_model=HealthResponse)
async def health_check():
    """
    Health check endpoint for monitoring.

    Returns 200 if healthy, 503 if unhealthy.
    """
    template_available = TEMPLATE_PATH.exists()
    timestamp = datetime.now(timezone.utc).isoformat()

    if template_available:
        return HealthResponse(
            status="healthy",
            version=API_VERSION,
            template_available=True,
            timestamp=timestamp,
        )
    else:
        return JSONResponse(
            status_code=503,
            content=HealthResponse(
                status="unhealthy",
                version=API_VERSION,
                template_available=False,
                timestamp=timestamp,
                error="Invoice template not found",
            ).model_dump(),
        )
