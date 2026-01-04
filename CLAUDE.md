# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

CCI Timesheet Reporting System - Python automation for generating weekly/monthly timesheet reports from Microsoft 365 calendars. Integrates with MS Graph API to collect time-tracking events marked with "XXX TIME CARD" naming convention, validates against project codes, and generates Apple Numbers reports and Excel invoices.

## Project Structure

```
cci-design/
├── src/                      # Main package
│   ├── core/                 # Core utilities
│   │   ├── config.py         # Constants, env vars, paths
│   │   ├── graph_client.py   # MS Graph client singleton
│   │   ├── database.py       # SQLite operations
│   │   └── validation.py     # Event validation rules
│   ├── services/             # Business logic
│   │   ├── calendar.py       # Calendar discovery & fetching
│   │   ├── reports.py        # Numbers report generation
│   │   ├── invoices.py       # Excel invoice generation
│   │   └── email.py          # Email sending
│   ├── models/               # Data models
│   │   └── events.py         # Event/Report TypedDicts
│   └── api/                  # FastAPI REST API
│       ├── main.py           # FastAPI app entry point
│       ├── dependencies.py   # API key auth dependency
│       ├── logging.py        # SQLite request logging
│       ├── routes/           # API route handlers
│       │   ├── health.py     # GET /health endpoint
│       │   └── invoices.py   # POST /v1/invoices/generate endpoint
│       └── models/
│           └── responses.py  # Pydantic response models
│   └── scripts/              # CLI entry points
│       ├── create_weekly_report.py
│       ├── create_monthly_report.py
│       ├── create_invoices.py
│       └── init_db.py
├── tests/                    # Test suite
│   ├── fixtures/             # Test data generators
│   ├── unit/                 # Unit tests
│   └── integration/          # Integration tests
├── data/
│   ├── db/                   # SQLite database
│   ├── reference/            # Reference spreadsheets
│   └── templates/            # Invoice templates
└── output/                   # Generated files (gitignored)
    ├── reports/weekly/
    ├── reports/monthly/
    └── invoices/
```

## Development Commands

```bash
# Initialize/reset database
uv run python src/scripts/init_db.py

# Generate test data (November 2025 synthetic events)
uv run python tests/fixtures/generate_events.py

# Sync test data to MS365 calendar
uv run python tests/fixtures/sync_to_calendar.py

# Run weekly report (defaults to today)
uv run python src/scripts/create_weekly_report.py

# Run weekly report for specific date
uv run python src/scripts/create_weekly_report.py --date 2025-11-07

# Run monthly report (defaults to previous month)
uv run python src/scripts/create_monthly_report.py

# Run monthly report for specific month
uv run python src/scripts/create_monthly_report.py --month 2025-11

# Generate invoices from monthly report
uv run python src/scripts/create_invoices.py output/reports/monthly/timesheet_monthly_report_2025_11_a.numbers

# Run API server (development)
API_DEBUG=true uv run uvicorn src.api.main:app --reload --host 0.0.0.0 --port 8000

# Run API server (production)
uv run uvicorn src.api.main:app --host 0.0.0.0 --port 8000

# Run tests
uv run pytest
```

## Architecture

### Core Modules (`src/core/`)
- **config.py**: Configuration constants, paths, validation codes
- **graph_client.py**: MS Graph client with lazy initialization
- **database.py**: SQLite operations for reports and events
- **validation.py**: Multi-stage event validation logic

### Services (`src/services/`)
- **calendar.py**: Calendar discovery and event fetching from MS Graph
- **reports.py**: Numbers file generation utilities
- **invoices.py**: Excel invoice generation from timesheet data
- **email.py**: Email formatting and sending via MS Graph

### Data Flow
1. Async scan of org users for "XXX TIME CARD" calendars via MS Graph
2. Fetch calendar events within date range
3. Parse event body for project_id, task, phase, wid
4. Multi-stage validation (presence → code validity → project consistency)
5. Generate Numbers file:
   - Weekly: Summary (pivot) + Detail sheets
   - Monthly: Two Detail sheets (View + Edit with identical data)
6. Email report with conflict summary
7. Generate invoices from monthly reports (Excel format)

### Database Schema (`data/db/cci-timesheets.db`)
- `reports`: Tracks each report run (id, type, name, create_date)
- `events`: Event data linked to reports (report_id FK, project_id, employee_id, timestamps, task, phase, wid, error_message)

## Validation Rules

### Valid Codes
- **Task codes**: BD, DP, PM, 3-D, D-D, M, NA
- **Phase codes**: PD, SD, DD, CD, CA, M, NA

### Project Types (determined by name prefix before `:`)
- **Non-projects** (office, vacation, holiday, sick, personal time):
  - Office: Task=BD or NA, Phase from {NA, PD, SD}
  - Others: Task=NA, Phase=NA
- **Regular projects**: Task from {BD, DP, PM, 3-D, D-D, M} (not NA), Phase from {PD, SD, DD, CD, CA, M} (not NA), cannot use BD task
- **Meetings**: Always Task=M, Phase=M

### Consistency Rule
Same project name cannot have multiple Project IDs in the same month.

## Environment Setup

Required `.env` variables:
```
# Azure/Microsoft Entra authentication
MICROSOFT_GRAPH_TENANT_ID=
MICROSOFT_GRAPH_APP_ID=
MICROSOFT_GRAPH_CLIENT_SECRET=

# API authentication (for FastAPI)
CCI_API_KEY=
```

## Key Patterns

- **Async operations**: All MS Graph API calls use asyncio for concurrent fetching
- **Timezone handling**: Events stored as UTC ISO8601; conversions use `ZoneInfo("EST"/"EDT")`
- **Report naming**: `{type}_report_{YYYY}_{MM}_{DD}_{suffix}` with auto-incrementing suffix (a, b, c...)
- **Numbers-parser**: Uses stable write support (v3.4.0+); creates docs with Summary (pivot) and Detail sheets
- **Service layer**: Business logic in `src/services/` can be used by CLI scripts or FastAPI endpoints

## REST API

FastAPI-based REST API for integration with Apple Shortcuts and other clients.

### Configuration
- **Framework**: FastAPI with uvicorn ASGI server
- **Authentication**: API key verification via `X-API-Key` header
- **API Key**: Stored in `.env` as `CCI_API_KEY`

### Endpoints

| Method | Endpoint | Auth | Description |
|--------|----------|------|-------------|
| GET | `/health` | No | Health check with template availability |
| POST | `/v1/invoices/generate` | Yes | Generate Excel invoices from Numbers file |

### Architecture
```
Apple Shortcuts → FastAPI (src/api/) → Services (src/services/) → Output
```

The API imports and calls the same service functions used by CLI scripts, ensuring consistent behavior.

### API Modules (`src/api/`)
- **main.py**: FastAPI app with lifespan handler, CORS middleware (debug mode), exception handlers
- **dependencies.py**: API key verification using constant-time comparison
- **logging.py**: SQLite request logging with `api_requests` and `api_request_details` tables
- **routes/health.py**: Health check endpoint returning template status
- **routes/invoices.py**: Invoice generation endpoint with file upload handling
- **models/responses.py**: Pydantic models (HealthResponse, ErrorResponse, ErrorCodes)

### Request Flow
1. File uploaded as multipart form-data
2. API key verified via `X-API-Key` header
3. File written to temp file (numbers-parser requires file path)
4. Processing runs in thread pool via `asyncio.to_thread()`
5. Excel bytes returned with Content-Disposition header
6. Request logged to SQLite database

### Database Tables (API Logging)
- `api_requests`: Request metadata (endpoint, status, timing, file info)
- `api_request_details`: Validation errors and warnings per request
