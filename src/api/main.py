"""FastAPI application entry point."""

import warnings
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from api.models.responses import ErrorCodes, ErrorResponse
from api.routes import health_router, invoices_router
from core.config import API_DEBUG, API_VERSION


@asynccontextmanager
async def lifespan(app: FastAPI):
    """Application lifespan handler for startup/shutdown."""
    # Startup: verify critical paths exist
    from services.invoices import TEMPLATE_PATH

    if not TEMPLATE_PATH.exists():
        warnings.warn(f"Invoice template not found at {TEMPLATE_PATH}")

    yield

    # Shutdown: cleanup if needed
    pass


app = FastAPI(
    title="CCI Invoice Generation API",
    description="REST API for generating Excel invoices from Apple Numbers monthly timesheet reports",
    version=API_VERSION,
    debug=API_DEBUG,
    lifespan=lifespan,
)

# CORS middleware (for development)
if API_DEBUG:
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )


# Global exception handler for unexpected errors
@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    """Handle unexpected exceptions with standard error format."""
    return JSONResponse(
        status_code=500,
        content=ErrorResponse(
            error="Internal server error",
            code=ErrorCodes.INTERNAL_ERROR,
            details=[],
        ).model_dump(),
    )


# Include routers
app.include_router(health_router)
app.include_router(invoices_router)


# Entry point for uvicorn
if __name__ == "__main__":
    import uvicorn

    from core.config import API_HOST, API_PORT

    uvicorn.run(
        "api.main:app",
        host=API_HOST,
        port=API_PORT,
        reload=API_DEBUG,
    )
