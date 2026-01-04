"""FastAPI dependencies for authentication and shared resources."""

import secrets

from fastapi import Header, HTTPException, status

from core.config import CCI_API_KEY


async def verify_api_key(x_api_key: str = Header(..., alias="X-API-Key")) -> str:
    """
    Verify API key from X-API-Key header.

    Raises:
        HTTPException: 401 if key is missing or invalid
    """
    if not CCI_API_KEY:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail={
                "error": "API key not configured on server",
                "code": "INTERNAL_ERROR",
                "details": [],
            },
        )

    # Use constant-time comparison to prevent timing attacks
    if not secrets.compare_digest(x_api_key, CCI_API_KEY):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail={
                "error": "Invalid or missing API key",
                "code": "UNAUTHORIZED",
                "details": [],
            },
        )

    return x_api_key
