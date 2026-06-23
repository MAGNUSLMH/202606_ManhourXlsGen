from __future__ import annotations

from typing import List, Optional

import xlsxwriter
from openpyxl import load_workbook


def _is_alphanumeric(value) -> bool:
    """Return True if value is a non-empty string of only A–Z / a–z / 0–9.

    Used to detect working order codes in the timesheet columns.
    """

    if value is None:
        return False
    s = str(value).strip()
    return bool(s) and s.isalnum()


def build_workbook(
    filename: str = "sample.xlsx",
    staffID: str = "",
    dates: Optional[List[int]] = None,
    *,
    timesheet_path: str,
    jCol1: str,
    jCol2: str,
    nCol1: str,
    nCol2: str,
    oCol1: str,
    oCol2: str,
) -> None:
    """Create an Excel workbook with a styled header row in A1:L1 and data rows.

    The header row has:
    - Grey background fill
    - Thin border around each cell
    - Bold, underlined text
    - Red font for specific headings, black for the rest
    - Autofilter drop-downs on each header cell (A1:L1)

    For each working order discovered in the source timesheet, one record
    row is created in the output. Columns A, B, C are always written as
    ``"A"``, ``"N"`` and the provided ``staffID`` respectively.
    """

    # Ensure we always have an iterable list for dates.
    if dates is None:
        dates = []

    # Open the source timesheet from which dates and working orders are read.
    # Use the sheet named "Timesheet" as requested, falling back to the
    # active sheet only if it does not exist.
    src_wb = load_workbook(timesheet_path, data_only=True)
    try:
        src_ws = src_wb["Timesheet"]
    except KeyError:
        src_ws = src_wb.active

    # Create the workbook and a single worksheet for the output.
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # Common formatting properties for all header cells.
    base_header_props = {
        "bold": True,
        "font_name": "Arial",  # header font family
        "underline": 1,  # single underline
        "bg_color": "#D9D9D9",  # light grey background similar to screenshot
        "border": 1,  # thin border on all sides
        "align": "center",
        "valign": "vcenter",
    }

    # Define two formats: one with red font and one with black font.
    red_header_format = workbook.add_format({
        **base_header_props,
        "font_color": "#C00000",  # dark red, close to Excel's standard
    })

    black_header_format = workbook.add_format({
        **base_header_props,
        "font_color": "#000000",
    })

    # Date format used for the TransDate column (column G).
    date_format = workbook.add_format({"num_format": "d/m/yyyy"})

    # Define the header labels and which color they should use.
    # (Based on the provided screenshot.)
    headers = [
        ("Action", "red"),
        ("AdjInd", "red"), # AdjInd, fixed
        ("StaffID", "red"),
        ("GlobalTemp", "black"),
        ("WorkOrderNum", "black"),
        ("ProjectNo", "black"),
        ("TransDate", "red"),
        ("Manhour", "red"),
        ("TransType", "red"),
        ("F1", "black"),
        ("ABCCode", "black"),
        ("TransRemark", "red"),
    ]

    # Write the header labels to row 1 (index 0) with appropriate formatting.
    # This recreates the styled header row shown in the screenshot.
    for col_idx, (label, color) in enumerate(headers):
        cell_format = red_header_format if color == "red" else black_header_format
        worksheet.write(0, col_idx, label, cell_format)

    # Optionally adjust column widths so that the text is visible and similar
    # to the appearance in the screenshot. This is not strictly required
    # for functionality but makes the result easier to inspect.
    worksheet.set_column("A:A", 9)   # Action
    worksheet.set_column("B:B", 9)   # AdjInc
    worksheet.set_column("C:C", 10)  # StaffID
    worksheet.set_column("D:D", 12)  # GlobalTemp
    worksheet.set_column("E:E", 14)  # WorkOrderNum
    worksheet.set_column("F:F", 11)  # ProjectNo
    worksheet.set_column("G:G", 11, date_format)  # TransDate with d/m/yyyy
    worksheet.set_column("H:H", 10)  # Manhour
    worksheet.set_column("I:I", 11)  # TransType
    worksheet.set_column("J:J", 5)   # F1
    worksheet.set_column("K:K", 9)   # ABCCode
    worksheet.set_column("L:L", 13)  # TransRemark

    # Format for data cells in columns E and H (thin borders around filled
    # cells, as required).
    data_border_format = workbook.add_format({"border": 1})

    # Data rows start at Excel row 2 (zero-based index 1 in xlsxwriter).
    data_start_row = 1
    current_row = data_start_row
    records_written = 0

    def _process_date_block(date_row: int, date_value, work_col: str, hours_col: str, ot_col: str) -> None:
        """Scan downward from a given date row for working orders in *work_col*.

        For every qualifying working order row we emit one record into the
        output sheet, using ``date_value`` as the TransDate and the matching
        hours / OT cells from ``hours_col`` and ``ot_col``.
        """

        nonlocal current_row, records_written

        row = date_row + 7
        while row <= src_ws.max_row:
            work_val = src_ws[f"{work_col}{row}"].value
            hours_val = src_ws[f"{hours_col}{row}"].value
            ot_val = src_ws[f"{ot_col}{row}"].value

            # If an entire trio of cells is empty we allow for a limited number
            # of blank rows; after that we stop scanning this block.
            if work_val is None and hours_val is None and ot_val is None:
                if row >= date_row + 52:
                    break
                row += 1
                continue

            if _is_alphanumeric(work_val) and _is_alphanumeric(hours_val) and str(work_val) != "0" and str(hours_val) != "0":
                # Columns A, B, C
                worksheet.write(current_row, 0, "A")
                worksheet.write(current_row, 1, "N")
                worksheet.write(current_row, 2, staffID)

                # Column E: WorkOrderNum (with thin border)
                worksheet.write(current_row, 4, work_val, data_border_format)

                # Column G: TransDate
                worksheet.write(current_row, 6, date_value)

                # Column H: Manhour (nCol value). Non-empty cells should
                # have thin borders.
                if hours_val not in (None, ""):
                    worksheet.write(current_row, 7, hours_val, data_border_format)

                # Column I: TransType ("OT" if there is a value in the OT
                # column, otherwise left blank). Column I cells do *not*
                # have borders.
                trans_type = "OT" if ot_val not in (None, "") else ""
                if trans_type:
                    worksheet.write(current_row, 8, trans_type)

                current_row += 1
                records_written += 1

            row += 1

    # Walk through the supplied list of candidate date rows. For each row we
    # process the date in column C using column set 1, then the date in column
    # J using column set 2. Both cells are assumed to contain dates.
    for date_row in dates:
        c_val = src_ws[f"C{date_row}"].value
        if c_val not in (None, ""):
            _process_date_block(date_row, c_val, jCol1, nCol1, oCol1)

        j_val = src_ws[f"J{date_row}"].value
        if j_val not in (None, ""):
            _process_date_block(date_row, j_val, jCol2, nCol2, oCol2)

    # Add an autofilter that covers the header row and all data rows.
    # Using zero-based (row, col) indices here.
    last_filter_row = data_start_row + records_written if records_written else 0
    worksheet.autofilter(0, 0, last_filter_row, 11)

    # Finalize and save the workbook.
    workbook.close()
    src_wb.close()


def main() -> None:
    """Entry point when running this script directly."""

    # Staff ID used for every record written to columns A, B, C.
    staffID = "309605"  # TODO: replace with the actual staff id as needed.

    # Alphabetic column identifiers for the two column sets used when
    # extracting working orders and hours from the source timesheet.
    #
    # Column set 1 is used when the date is in column C (C4, C63, ...):
    #   - jCol1: working order codes
    #   - nCol1: normal hours
    #   - oCol1: OT hours
    # Column set 2 is used when the date is in column J (J4, J63, ...):
    #   - jCol2: working order codes
    #   - nCol2: normal hours
    #   - oCol2: OT hours
    # Adjust these letters if your timesheet layout changes.
    jCol1, jCol2 = "A", "I"
    nCol1, nCol2 = "C", "J"
    oCol1, oCol2 = "D", "K"
    # Fixed list of candidate date rows in the source timesheet. The
    # workbook logic will check C4, J4, C63, J63, ... in this order until
    # an empty pair is encountered.
    dates = list(range(4, 949, 59))  # This number list is fixed

    # Path to the source timesheet workbook.
    timesheet_path = "Timesheet Record Team A_2026 May.xlsx"

    # Create the output workbook using the configured parameters.
    build_workbook(
        "sample.xlsx",
        staffID=staffID,
        dates=dates,
        timesheet_path=timesheet_path,
        jCol1=jCol1,
        jCol2=jCol2,
        nCol1=nCol1,
        nCol2=nCol2,
        oCol1=oCol1,
        oCol2=oCol2,
    )


if __name__ == "__main__":  # pragma: no cover - manual execution guard
    main()
