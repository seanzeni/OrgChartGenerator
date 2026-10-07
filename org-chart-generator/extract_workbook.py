import argparse
import json
import sys
from pathlib import Path

from openpyxl import load_workbook


def normalize(value):
    if value is None:
        return ""
    return str(value).strip()


def read_sheet(workbook, sheet_name, columns):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Sheet '{sheet_name}' was not found. Available sheets: {', '.join(workbook.sheetnames)}"
        )

    sheet = workbook[sheet_name]
    rows = sheet.iter_rows(values_only=True)
    try:
        headers = [normalize(value) for value in next(rows)]
    except StopIteration:
        raise ValueError(f"Sheet '{sheet_name}' is empty")

    header_lookup = {header.casefold(): index for index, header in enumerate(headers) if header}
    missing = [label for label in columns.values() if label.casefold() not in header_lookup]
    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing columns: {', '.join(missing)}. "
            f"Found: {', '.join(headers)}"
        )

    records = []
    for row_number, row in enumerate(rows, start=2):
        record = {}
        for key, label in columns.items():
            index = header_lookup[label.casefold()]
            record[key] = normalize(row[index] if index < len(row) else None)
        if not any(record.values()):
            continue
        if not record["employeeId"]:
            raise ValueError(f"Sheet '{sheet_name}', row {row_number}: employee ID is blank")
        if len(record["employeeId"]) != 4 or not record["employeeId"].isalnum():
            raise ValueError(
                f"Sheet '{sheet_name}', row {row_number}: employee ID "
                f"'{record['employeeId']}' must contain exactly four letters or numbers"
            )
        if not record["employeeName"]:
            raise ValueError(f"Sheet '{sheet_name}', row {row_number}: employee name is blank")
        if not record["managerId"]:
            raise ValueError(f"Sheet '{sheet_name}', row {row_number}: manager ID is blank")
        if len(record["managerId"]) != 4 or not record["managerId"].isalnum():
            raise ValueError(
                f"Sheet '{sheet_name}', row {row_number}: manager ID "
                f"'{record['managerId']}' must contain exactly four letters or numbers"
            )
        if not record["managerName"]:
            raise ValueError(f"Sheet '{sheet_name}', row {row_number}: manager name is blank")
        if not record["role"]:
            record["role"] = "Unspecified"
        record["sourceRow"] = row_number
        records.append(record)

    duplicates = {}
    for record in records:
        key = record["employeeId"].casefold()
        duplicates.setdefault(key, []).append(record["sourceRow"])
    duplicate_text = [f"{key}: rows {', '.join(map(str, rows))}" for key, rows in duplicates.items() if len(rows) > 1]
    if duplicate_text:
        raise ValueError(f"Sheet '{sheet_name}' has duplicate employee IDs: {'; '.join(duplicate_text)}")
    return records


def main():
    parser = argparse.ArgumentParser(description="Extract two organization sheets to JSON")
    parser.add_argument("--settings", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()

    settings_path = Path(args.settings).resolve()
    settings = json.loads(settings_path.read_text(encoding="utf-8"))
    workbook_path = Path(settings["workbookPath"])
    if not workbook_path.is_absolute():
        workbook_path = (settings_path.parent / workbook_path).resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    payload = {
        "workbookPath": str(workbook_path),
        "current": read_sheet(workbook, settings["sheets"]["current"], settings["columns"]),
        "future": read_sheet(workbook, settings["sheets"]["future"], settings["columns"]),
    }
    Path(args.output).write_text(json.dumps(payload, indent=2), encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)
