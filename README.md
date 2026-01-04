# CCI Timesheet Reporting System

Python automation for generating weekly/monthly timesheet reports from Microsoft 365 calendars, with Excel invoice generation.

## Features

- **Calendar Integration**: Scans Microsoft 365 calendars for time-tracking events using the "XXX TIME CARD" naming convention
- **Weekly Reports**: Generates Apple Numbers files with Summary (pivot table) and Detail sheets
- **Monthly Reports**: Creates Numbers files with View and Edit detail sheets
- **Invoice Generation**: Produces professionally formatted Excel invoices from monthly timesheet data
- **Validation**: Multi-stage validation of task/phase codes with conflict detection
- **Email Notifications**: Sends reports via email with conflict summaries

## Requirements

- Python 3.12+
- [uv](https://github.com/astral-sh/uv) package manager
- Microsoft 365 account with Graph API access
- Azure/Entra app registration

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd cci-design

# Install dependencies
uv sync
```

## Configuration

Create a `.env` file in the project root:

```env
MICROSOFT_GRAPH_TENANT_ID=your-tenant-id
MICROSOFT_GRAPH_APP_ID=your-app-id
MICROSOFT_GRAPH_CLIENT_SECRET=your-client-secret

# API authentication
CCI_API_KEY=your-secret-api-key
```

## Usage

### Initialize Database

```bash
uv run python src/scripts/init_db.py
```

### Generate Weekly Report

```bash
# Defaults to today
uv run python src/scripts/create_weekly_report.py

# Specific date
uv run python src/scripts/create_weekly_report.py --date 2025-11-07
```

### Generate Monthly Report

```bash
# Defaults to previous month
uv run python src/scripts/create_monthly_report.py

# Specific month
uv run python src/scripts/create_monthly_report.py --month 2025-11
```

### Generate Invoices

```bash
uv run python src/scripts/create_invoices.py output/reports/monthly/timesheet_monthly_report_2025_11_a.numbers
```

### Run API Server

```bash
# Development (with auto-reload)
API_DEBUG=true uv run uvicorn src.api.main:app --reload --host 0.0.0.0 --port 8000

# Production
uv run uvicorn src.api.main:app --host 0.0.0.0 --port 8000
```

## REST API

The API provides programmatic access to invoice generation, designed for integration with Apple Shortcuts and other clients.

### Endpoints

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/health` | Health check with template status |
| POST | `/v1/invoices/generate` | Generate invoices from Numbers file |

### Authentication

All `/v1/*` endpoints require an API key via the `X-API-Key` header.

### Generate Invoices via API

```bash
curl -X POST http://localhost:8000/v1/invoices/generate \
  -H "X-API-Key: your-api-key" \
  -F "file=@timesheet_monthly_report_2025_11_a.numbers" \
  -o invoices.xlsx
```

Optional parameter:
- `invoice_date`: Override invoice date (YYYY-MM-DD format)

## Project Structure

```
cci-design/
├── src/
│   ├── core/           # Config, database, validation, Graph client
│   ├── services/       # Calendar, reports, invoices, email
│   ├── models/         # Data models
│   ├── api/            # FastAPI REST API
│   └── scripts/        # CLI entry points
├── tests/              # Test suite
├── data/
│   ├── db/             # SQLite database
│   ├── reference/      # Reference spreadsheets
│   └── templates/      # Invoice templates
└── output/             # Generated reports and invoices
```

## Calendar Event Format

Events must follow this naming convention:
- Calendar name: `XXX TIME CARD` (e.g., "CES TIME CARD")
- Event subject: `Project Name: Project Number`
- Event body must include:
  ```
  WID: Work item description
  Task: PM
  Phase: SD
  ```

### Valid Codes

| Task Codes | Phase Codes |
|------------|-------------|
| BD, DP, PM, 3-D, D-D, M, NA | PD, SD, DD, CD, CA, M, NA |

## Output

- **Weekly Reports**: `output/reports/weekly/timesheet_weekly_report_YYYY_MM_DD_a.numbers`
- **Monthly Reports**: `output/reports/monthly/monthly_report_YYYY_MM_DD_a.numbers`
- **Invoices**: `output/invoices/invoices_month_year_a.xlsx`

## Development

```bash
# Run tests
uv run pytest

# Generate test data
uv run python tests/fixtures/generate_events.py
```

## License

Private - CCI Design Inc.
