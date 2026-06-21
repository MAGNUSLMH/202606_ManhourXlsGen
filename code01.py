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
    src_wb = load_workbook(timesheet_path, data_only=True)
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

    # Define the header labels and which color they should use.
    # (Based on the provided screenshot.)
    headers = [
        ("Action", "red"),
        ("AdjInc", "red"),
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

    # Optionally adjust column widths so that the text is visible and similar
    # to the appearance in the screenshot. This is not strictly required
    # for functionality but makes the result easier to inspect.
    worksheet.set_column("A:A", 9)   # Action
    worksheet.set_column("B:B", 9)   # AdjInc
    worksheet.set_column("C:C", 10)  # StaffID
    worksheet.set_column("D:D", 12)  # GlobalTemp
    worksheet.set_column("E:E", 14)  # WorkOrderNum
    worksheet.set_column("F:F", 11)  # ProjectNo
    worksheet.set_column("G:G", 11)  # TransDate
    worksheet.set_column("H:H", 10)  # Manhour
    worksheet.set_column("I:I", 11)  # TransType
    worksheet.set_column("J:J", 5)   # F1
    worksheet.set_column("K:K", 9)   # ABCCode
    worksheet.set_column("L:L", 13)  # TransRemark

    # Write the header labels to row 1 (index 0) with appropriate formatting.
    for col_idx, (label, color) in enumerate(headers):
        cell_format = red_header_format if color == "red" else black_header_format
        worksheet.write(0, col_idx, label, cell_format)

    # Format for data cells in columns E and I (thin borders around filled
    # cells, as required).
    data_border_format = workbook.add_format({"border": 1})

    # Data rows start at Excel row 2 (zero-based index 1 in xlsxwriter).
    data_start_row = 1
    current_row = data_start_row
    records_written = 0

    # Walk through the supplied list of candidate date rows. For each row
    # we first look at column C; if it contains a date, we use column set 1
    # (jCol1, nCol1, oCol1). Otherwise we look at column J and, if that
    # contains a date, we use column set 2. If both cells are empty we
    # stop processing any further rows.
    processed_first_date = False

    for date_row in dates:
        c_cell_ref = f"C{date_row}"
        j_cell_ref = f"J{date_row}"
        c_val = src_ws[c_cell_ref].value

        # Default to column set 2 (J-based) unless C has the date.
        if c_val is not None and c_val != "":
            date_value = c_val
            work_col = jCol1
            hours_col = nCol1
            ot_col = oCol1
        else:
            j_val = src_ws[j_cell_ref].value
            if j_val is None or j_val == "":
                # Both C and J are empty for this row: terminate the
                # process as requested.
                break

            date_value = j_val
            work_col = jCol2
            hours_col = nCol2
            ot_col = oCol2

        # Scan downward in the chosen working-order column to find all
        # alphanumeric work order codes. For each such code we create one
        # record row in the output sheet.
        row = date_row + 1
        while row <= src_ws.max_row:
            work_val = src_ws[f"{work_col}{row}"].value
            hours_val = src_ws[f"{hours_col}{row}"].value
            ot_val = src_ws[f"{ot_col}{row}"].value

            # If the entire trio is empty we assume the block has ended.
            if work_val is None and hours_val is None and ot_val is None:
                break

            if _is_alphanumeric(work_val):
                # Columns A, B, C
                worksheet.write(current_row, 0, "A")
                worksheet.write(current_row, 1, "N")
                worksheet.write(current_row, 2, staffID)

                # Column E: WorkOrderNum (with thin border)
                worksheet.write(current_row, 4, work_val, data_border_format)

                # Column G: TransDate
                worksheet.write(current_row, 6, date_value)

                # Column H: Manhour (nCol value)
                worksheet.write(current_row, 7, hours_val)

                # Column I: TransType ("OT" if there is a value in the OT
                # column, otherwise left blank). Filled cells get borders.
                trans_type = "OT" if ot_val not in (None, "") else ""
                if trans_type:
                    worksheet.write(current_row, 8, trans_type, data_border_format)

                current_row += 1
                records_written += 1

            row += 1

        # Temporary break so that, for now, only the first date row
        # (e.g. C4 / J4) is processed.
        processed_first_date = True
        if processed_first_date:
            break

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
    staffID = "STAFF_ID"  # TODO: replace with the actual staff id as needed.

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
    dates = [4, 63]  # Extend this list as needed.

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
