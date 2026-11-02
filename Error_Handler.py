import csv
import math
from datetime import date, datetime
from pathlib import Path

import openpyxl #xlsx, xlsm
import xlrd #xls


script_dir = Path(__file__).resolve().parent
supported_extensions = (".csv", ".xls", ".xlsx", ".xlsm")


def resolve_source_file(entered_name):
    name = entered_name.strip().strip('"')
    if not name or Path(name).name != name:
        raise FileNotFoundError(
            "File cannot be found. Check if it's the right extension "
            f"({', '.join(supported_extensions)})."
        )
    suffix = Path(name).suffix.casefold()
    if suffix:
        if suffix not in supported_extensions:
            raise FileNotFoundError(
                "File cannot be found. Check if it's the right extension "
                f"({', '.join(supported_extensions)})."
            )
        candidates = [script_dir / name]
    else:
        candidates = [
            script_dir / f"{name}{extension}"
            for extension in (".xls", ".xlsx")
        ]
    existing = [path for path in candidates if path.is_file()]
    if not existing:
        raise FileNotFoundError(
            "File cannot be found. Check if it's the right extension "
            f"({', '.join(supported_extensions)})."
        )
    if len(existing) > 1:
        matches = ", ".join(path.name for path in existing)
        raise ValueError(
            f"Multiple source files match {name!r}: {matches}. Specify the extension."
        )
    return existing[0]


class CellValue:

    def __init__(self, value):
        self.value = value
        self.ctype = (
            xlrd.XL_CELL_DATE
            if isinstance(value, (date, datetime))
            else xlrd.XL_CELL_EMPTY
            if value is None
            else xlrd.XL_CELL_TEXT
        )


class WorksheetValues:

    def __init__(self, rows):
        self._rows = [list(row) for row in rows]
        self.nrows = len(self._rows)
        self.ncols = max((len(row) for row in self._rows), default=0)

    def row_values(self, row_index):
        row = self._rows[row_index]
        return row + [""] * (self.ncols - len(row))

    def cell_value(self, row_index, column_index):
        if column_index >= len(self._rows[row_index]):
            return ""
        value = self._rows[row_index][column_index]
        return "" if value is None else value

    def cell(self, row_index, column_index):
        return CellValue(self.cell_value(row_index, column_index))


class SourceWorkbook:
    def __init__(self, worksheet, datemode=0):
        self._worksheet = worksheet
        self.datemode = datemode

    def sheet_by_index(self, index):
        if index != 0:
            raise IndexError("The source file contains only one worksheet.")
        return self._worksheet


def open_source_workbook(path):
    suffix = path.suffix.casefold()
    if suffix == ".xls":
        workbook = xlrd.open_workbook(str(path))
        return SourceWorkbook(workbook.sheet_by_index(0), workbook.datemode)
    if suffix in (".xlsx", ".xlsm"):
        workbook = openpyxl.load_workbook(path, data_only=True, read_only=True)
        try:
            worksheet = next(iter(workbook.worksheets))
            rows = WorksheetValues(worksheet.iter_rows(values_only=True))
        finally:
            workbook.close()
        return SourceWorkbook(rows)
    if suffix == ".csv":
        with path.open(newline="", encoding="utf-8-sig") as csv_file:
            return SourceWorkbook(WorksheetValues(csv.reader(csv_file)))
    raise ValueError(f"Unsupported source file extension: {suffix}")


source_workbook_filename = input(
    "Enter the source filename (extension optional): "
)
workbook_path = resolve_source_file(source_workbook_filename)
source_workbook_filename = workbook_path.name
workbook_name = Path(source_workbook_filename).stem
error_dir = script_dir / f"Error Files_{workbook_name}"
resolved_dir = script_dir / f"Resolved Files_{workbook_name}"
date_formats = (
    "%Y/%m/%d %H:%M:%S",
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S",
    "%Y/%m/%d %H:%M",
    "%Y-%m-%d %H:%M",
    "%y/%m/%d %H:%M:%S",
    "%Y/%m/%d",
)


def normalize(value):
    return " ".join(str(value).split()).casefold()


def is_numeric(value):
    try:
        return math.isfinite(float(value.strip()))
    except ValueError:
        return False


def valid_serial_number(value):
    if isinstance(value, float) and value.is_integer():
        serial_number = str(int(value))
    else:
        serial_number = str(value).strip()
    if (
        len(serial_number) == 12
        and serial_number.startswith("19")
        and all("0" <= digit <= "9" for digit in serial_number)
    ):
        return serial_number
    return None


def parse_timestamp(cell, datemode):
    if isinstance(cell.value, datetime):
        return cell.value
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode)
    value = " ".join(str(cell.value).split())
    for date_format in date_formats:
        try:
            return datetime.strptime(value, date_format)
        except ValueError:
            continue
    return None

# Judge the other test results, not the verdict being recalculated or version fields
def assess_csv_verdict(headers, row):
    excluded_fields = {"fpga version", "zigbee version", "overall verdict"}
    fail_values = {"", "n/a", "fail", "error"}
    for header, value in zip(headers, row):
        if normalize(header) not in excluded_fields and value.strip().casefold() in fail_values:
            return "Fail"
    return "Pass"

# Match error CSV filenames to source rows and collect their verdict values
def find_verdict_mismatches(csv_files):
    if not workbook_path.is_file():
        raise FileNotFoundError(f"Excel workbook not found: {workbook_path}")

    workbook = open_source_workbook(workbook_path)
    worksheet = workbook.sheet_by_index(0)
    header_row = None
    source_columns = {}

    for row_index in range(worksheet.nrows):
        headers = [str(value).strip().casefold() for value in worksheet.row_values(row_index)]
        if (
            "time and date" in headers
            and "serial number" in headers
            and "overall verdict" in headers
        ):
            header_row = row_index
            source_columns = {
                "time": headers.index("time and date"),
                "serial": headers.index("serial number"),
                "overall_verdict": headers.index("overall verdict") + 1,
            }
            break
    if header_row is None:
        raise ValueError("Could not find the required source workbook headers.")

    first_data_row = next(
        (
            row_index
            for row_index in range(header_row + 1, worksheet.nrows)
            if any(
                str(worksheet.cell_value(row_index, source_columns[column])).strip()
                for column in ("time", "serial")
            )
        ),
        worksheet.nrows,
    )
    header_rows = range(header_row, first_data_row)
    pass_fail_header = next(
        (
            (candidate_row, column)
            for candidate_row in header_rows
            for column, value in enumerate(worksheet.row_values(candidate_row))
            if "pass or fail the test" in normalize(value)
        ),
        None,
    )
    if pass_fail_header is None:
        raise ValueError("Could not find the Pass or Fail the Test header.")
    # The first column holds the colour; the next column contains the Pass/Fail value
    source_columns["pass_fail_test"] = pass_fail_header[1] + 1

    file_names = {path.name for path in csv_files}
    mismatches = {}
    for row_index in range(first_data_row, worksheet.nrows):
        serial_number = valid_serial_number(
            worksheet.cell_value(row_index, source_columns["serial"])
        )
        if serial_number is None:
            continue
        timestamp = parse_timestamp(
            worksheet.cell(row_index, source_columns["time"]),
            workbook.datemode,
        )
        if timestamp is None:
            continue
        filename = f"FT3_{serial_number}_{timestamp:%Y-%m-%d_%H-%M-%S}.csv"
        if filename not in file_names:
            continue

        # Retain source values for the prompt to resolve invalid or conflicting verdicts
        pass_fail_result = normalize(
            worksheet.cell_value(row_index, source_columns["pass_fail_test"])
        )
        overall_verdict = normalize(
            worksheet.cell_value(row_index, source_columns["overall_verdict"])
        )
        if pass_fail_result not in ("pass", "fail") or pass_fail_result != overall_verdict:
            mismatch = (
                timestamp,
                pass_fail_result,
                overall_verdict,
            )
            if filename in mismatches and mismatches[filename] != mismatch:
                raise ValueError(
                    f"Multiple source rows with different verdicts match {filename!r}."
                )
            mismatches[filename] = mismatch

    return mismatches


def prompt_to_resolve_verdict_mismatches(mismatches):
    """Optionally correct mismatch CSVs and move successfully resolved files."""
    while True:
        response = input(
            "Resolve Pass/Fail and Overall Verdict Mismatch? (y/n) "
        ).strip().casefold()
        if response in ("n", "no"):
            return set()
        if response in ("y", "yes"):
            break
        print("The input is invalid. Please enter y or n.")

    csv_assessed_conflicts = 0
    resolved_count = 0
    resolved_files = set()
    for filename, (timestamp, pass_fail_result, overall_verdict) in mismatches.items():
        pass_fail_valid = pass_fail_result in ("pass", "fail")
        overall_valid = overall_verdict in ("pass", "fail")
        if not pass_fail_valid and not overall_valid:
            continue

        # Locate the CSV row by both serial and timestamp before changing its verdict
        csv_path = error_dir / filename
        with csv_path.open(newline="", encoding="utf-8") as csv_file:
            rows = list(csv.reader(csv_file))
        if not rows:
            raise ValueError(f"CSV file is empty: {csv_path}")
        try:
            verdict_column = rows[0].index("Overall Verdict")
            serial_column = rows[0].index("Serial Number")
            time_column = rows[0].index("Test Time")
        except ValueError as error:
            raise ValueError(f"CSV is missing a required column: {csv_path}") from error

        matching_rows = []
        for row in rows[1:]:
            if len(row) != len(rows[0]):
                raise ValueError(f"Unexpected CSV row width in {csv_path}")
            row_timestamp = None
            for date_format in date_formats:
                try:
                    row_timestamp = datetime.strptime(row[time_column].strip(), date_format)
                    break
                except ValueError:
                    continue
            if (
                row[serial_column].strip() == filename[4:16]
                and row_timestamp == timestamp
            ):
                matching_rows.append(row)
        if len(matching_rows) != 1:
            raise ValueError(
                f"Expected one CSV data row matching the source record in {csv_path}; "
                f"found {len(matching_rows)}."
            )

        if pass_fail_valid and overall_valid:
            replacement_verdict = assess_csv_verdict(
                rows[0], matching_rows[0]
            )
            csv_assessed_conflicts += 1
        else:
            replacement_verdict = (
                pass_fail_result if pass_fail_valid else overall_verdict
            ).title()

        resolved_path = resolved_dir / filename
        if resolved_path.exists():
            raise FileExistsError(
                f"Cannot move resolved CSV because the destination already exists: {resolved_path}"
            )
        resolved_dir.mkdir(exist_ok=True)
        matching_rows[0][verdict_column] = replacement_verdict
        with csv_path.open("w", newline="", encoding="utf-8") as csv_file:
            csv.writer(csv_file).writerows(rows)
        csv_path.rename(resolved_path)
        resolved_count += 1
        resolved_files.add(filename)

    print(f"Files resolved and moved to Resolved Files: {resolved_count}")
    print(f"Conflicting valid source verdict files assessed from CSV data: {csv_assessed_conflicts}")
    if resolved_count:
        print(f"Resolved CSV files moved to: {resolved_dir}")
    return resolved_files


def prompt_for_error_file_names(error_entry_files, verdict_mismatch_files, abvf_files):
    """List file names by error category only when the user requests detail."""
    while True:
        response = input("More information provided of files? (y/n) ").strip().casefold()
        if response in ("n", "no"):
            return
        if response in ("y", "yes"):
            break
        print("The input is invalid. Please enter y or n.")

    categories = (
        ("Files with ERROR entries", error_entry_files),
        ("Unresolved Overall Verdict mismatch files", verdict_mismatch_files),
        ("Files with inconsistent ABVF entries", abvf_files),
    )
    for label, filenames in categories:
        print(f"{label}:")
        if filenames:
            for filename in sorted(filenames):
                print(f"  {filename}")
        else:
            print("  None")


if not error_dir.is_dir():
    print(f"No files found in {error_dir.name} folder.")
else:
    csv_files = sorted(
        path
        for path in error_dir.iterdir()
        if path.is_file() and path.suffix.casefold() == ".csv"
    )
    if not csv_files:
        print(f"No files found in {error_dir.name} folder.")
    else:
        files_with_error_entries = set()
        files_with_inconsistent_abvf = set()
        for path in csv_files:
            with path.open(newline="", encoding="utf-8") as csv_file:
                rows = list(csv.reader(csv_file))
            if not rows:
                raise ValueError(f"CSV file is empty: {path}")
            abvf_columns = [
                rows[0].index(f"ABVF {number}")
                for number in range(1, 5)
                if f"ABVF {number}" in rows[0]
            ]
            if len(abvf_columns) != 4:
                raise ValueError(f"CSV is missing one or more ABVF columns: {path}")
            if any(
                cell.strip().casefold() == "error"
                for row in rows[1:]
                for cell in row
            ):
                files_with_error_entries.add(path.name)

            # Each record must have four numeric values or four N/A values.
            for row in rows[1:]:
                abvf_values = [row[column].strip() for column in abvf_columns]
                all_na = all(value.casefold() == "n/a" for value in abvf_values)
                all_numeric = all(is_numeric(value) for value in abvf_values)
                if not all_na and not all_numeric:
                    files_with_inconsistent_abvf.add(path.name)
                    break

        verdict_mismatches = find_verdict_mismatches(csv_files)
        files_with_verdict_mismatches = set(verdict_mismatches)

        # Files can belong to multiple categories, so category totals are independent
        files_with_neither_reason = (
            {path.name for path in csv_files}
            - files_with_error_entries
            - files_with_verdict_mismatches
        )

        print(f"Total CSV files in {error_dir.name} folder: {len(csv_files)}")
        print(f"Files with ERROR entries: {len(files_with_error_entries)}")
        print(f"Files with Pass/Fail and Overall Verdict mismatch: {len(files_with_verdict_mismatches)}")
        print(f"Files with neither reason identified: {len(files_with_neither_reason)}")
        print(f"Files with inconsistent ABVF entries: {len(files_with_inconsistent_abvf)}")
        resolved_verdict_files = set()
        if verdict_mismatches:
            resolved_verdict_files = prompt_to_resolve_verdict_mismatches(verdict_mismatches)

        # Only report verdict mismatches still present after any requested resolutions
        remaining_csv_files = {
            path.name
            for path in error_dir.iterdir()
            if path.is_file() and path.suffix.casefold() == ".csv"
        }
        unresolved_verdict_files = (
            files_with_verdict_mismatches
            - resolved_verdict_files
        ) & remaining_csv_files
        prompt_for_error_file_names(
            files_with_error_entries & remaining_csv_files,
            unresolved_verdict_files,
            files_with_inconsistent_abvf & remaining_csv_files,
        )
