import sys
from numbers_parser import Document

def list_sheets_and_tables(filepath):
    """List all sheets and tables in a Numbers file."""
    doc = Document(filepath)

    print(f"File: {filepath}")
    print(f"Total sheets: {len(doc.sheets)}\n")

    for sheet_idx, sheet in enumerate(doc.sheets):
        print(f"Sheet {sheet_idx}: {sheet.name}")
        print(f"  Tables: {len(sheet.tables)}")

        for table_idx, table in enumerate(sheet.tables):
            num_rows = table.num_rows
            num_cols = table.num_cols
            print(f"    Table {table_idx}: {table.name} ({num_rows} rows x {num_cols} cols)")

        print()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python list_numbers.py <file.numbers>")
        sys.exit(1)

    list_sheets_and_tables(sys.argv[1])
