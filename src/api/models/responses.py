"""Pydantic response models for API endpoints."""

from pydantic import BaseModel


class HealthResponse(BaseModel):
    """Health check response."""

    status: str  # "healthy" or "unhealthy"
    version: str
    template_available: bool
    timestamp: str  # ISO 8601 UTC
    error: str | None = None


class ErrorResponse(BaseModel):
    """Standard error response."""

    error: str
    code: str
    details: list[str] = []


class ErrorCodes:
    """Error code constants."""

    INVALID_REQUEST = "INVALID_REQUEST"
    UNAUTHORIZED = "UNAUTHORIZED"
    FILE_TOO_LARGE = "FILE_TOO_LARGE"
    UNSUPPORTED_MEDIA_TYPE = "UNSUPPORTED_MEDIA_TYPE"
    VALIDATION_ERROR = "VALIDATION_ERROR"
    NO_BILLABLE_PROJECTS = "NO_BILLABLE_PROJECTS"
    INTERNAL_ERROR = "INTERNAL_ERROR"
