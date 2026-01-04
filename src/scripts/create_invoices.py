#!/usr/bin/env python3
"""
Generate client invoices from monthly timesheet reports.

Reads timesheet data from an Apple Numbers file, uses an Excel template to create
professionally formatted invoice cover sheets per project with Excel formulas,
and attaches detailed timesheet breakdowns.

Usage:
    uv run python src/scripts/create_invoices.py <input_file.numbers>

Example:
    uv run python src/scripts/create_invoices.py output/reports/monthly/timesheet_monthly_report_2025_11_a.numbers
"""

import argparse
import sys
from pathlib import Path

# Add src to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from services.invoices import generate_invoices


def main():
    parser = argparse.ArgumentParser(
        description="Generate client invoices from monthly timesheet reports"
    )
    parser.add_argument(
        "input_file",
        type=Path,
        help="Path to the monthly report Numbers file",
    )

    args = parser.parse_args()

    try:
        output_path = generate_invoices(args.input_file)
        print(f"\nInvoices generated: {output_path}")
    except Exception as e:
        print(f"\nError: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
