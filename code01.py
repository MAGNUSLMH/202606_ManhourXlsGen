from __future__ import annotations

from typing import List, Optional

from datetime import date, datetime

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


def _normalize_date(value):
    """Normalize various date/datetime representations to a date-like key.

    This helps us match dates between the main "Timesheet" sheet and the
    "WO" sheet when looking up shift codes.
    """

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return value


def _has_positive_hours(value) -> bool:
    """Return True if *value* represents a positive numeric hour value.

    This accepts integers, floats, or numeric strings such as "1.5" and is
    used instead of an alphanumeric check so that fractional hours (e.g.
    1.5) are not filtered out.
    """

    if value in (None, ""):
        return False

    try:
        return float(value) > 0
    except (TypeError, ValueError):
        return False


def build_workbook(
    filename: str = "sample.xlsx",
    staffID: Optional[List[str]] = None,
    dates: Optional[List[int]] = None,
    *,
    timesheet_path: str,
    jCol1: str,
    jCol2: str,
    nCol1: List[str],
    nCol2: List[str],
    oCol1: List[str],
    oCol2: List[str],
) -> None:
    """Create an Excel workbook with one worksheet per staff / column set.

    The header row on each worksheet has:
    - Grey background fill
    - Thin border around each cell
    - Bold, underlined text
    - Red font for specific headings, black for the rest
    - Autofilter drop-downs on each header cell (A1:L1)

    For each working order discovered in the source timesheet, one record
    row is created in the corresponding output worksheet. Columns A, B, C
    are always written as ``"A"``, ``"N"`` and the staff ID for that
    worksheet respectively.
    """

    # Ensure we always have iterable lists for dates and staff IDs.
    if dates is None:
        dates = []
    if staffID is None:
        staffID = []

    if not (len(staffID) == len(nCol1) == len(nCol2) == len(oCol1) == len(oCol2)):
        raise ValueError("staffID, nCol1, nCol2, oCol1, and oCol2 must all have the same length")

    # Open the source timesheet from which dates and working orders are read.
    # Use the sheet named "Timesheet" as requested, falling back to the
    # active sheet only if it does not exist.
    src_wb = load_workbook(timesheet_path, data_only=True)
    try:
        src_ws = src_wb["Timesheet"]
    except KeyError:
        src_ws = src_wb.active

    # Create the workbook for the output. Individual worksheets will be
    # created per staff / column-set combination.
    workbook = xlsxwriter.Workbook(filename)

    # Common formatting properties for all header cells.
    base_header_props = {
        "bold": True,
        "font_name": "Aptos",  # header font family as requested
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
    # Display format should be m/d/yyyy and centred.
    date_format = workbook.add_format({
        "num_format": "m/d/yyyy",
        "align": "center",
    })

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

    # Helper to set up a worksheet with header row and column widths.
    def _initialise_worksheet(ws) -> None:
        # Write the header labels to row 1 (index 0) with appropriate formatting.
        # This recreates the styled header row shown in the screenshot.
        for col_idx, (label, color) in enumerate(headers):
            cell_format = red_header_format if color == "red" else black_header_format
            ws.write(0, col_idx, label, cell_format)

        # Optionally adjust column widths so that the text is visible and similar
        # to the appearance in the screenshot. This is not strictly required
        # for functionality but makes the result easier to inspect.
        ws.set_column("A:A", 9)   # Action
        ws.set_column("B:B", 9)   # AdjInc
        ws.set_column("C:C", 10)  # StaffID
        ws.set_column("D:D", 12)  # GlobalTemp
        ws.set_column("E:E", 14)  # WorkOrderNum
        ws.set_column("F:F", 11)  # ProjectNo
        ws.set_column("G:G", 11, date_format)  # TransDate with m/d/yyyy
        ws.set_column("H:H", 10)  # Manhour
        ws.set_column("I:I", 11)  # TransType
        ws.set_column("J:J", 5)   # F1
        ws.set_column("K:K", 9)   # ABCCode
        ws.set_column("L:L", 13)  # TransRemark

    # Format for data cells in columns E and H (thin borders, centred) with
    # Aptos font, as requested.
    data_border_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "font_name": "Aptos",
    })

    # Format for StaffID data in column C (right-aligned as requested).
    staffid_right_format = workbook.add_format({"align": "right"})

    # Data rows start at Excel row 2 (zero-based index 1 in xlsxwriter).
    data_start_row = 1

    # Shift codes are now taken directly from the main timesheet sheet. For a
    # given date row N and normal-hours column (e.g. "C"), the shift code is
    # read from row N+3 in that same column (e.g. "C{N+3}").

    def _process_date_block(
        worksheet,
        staff_id: str,
        start_row: int,
        date_row: int,
        date_value,
        work_col: str,
        hours_col: str,
        ot_col: str,
        shift_code,
    ) -> int:
        """Scan downward from a given date row for working orders in *work_col*.

        For every qualifying working order row we emit one record into the
        output sheet, using ``date_value`` as the TransDate and the matching
        hours / OT cells from ``hours_col`` and ``ot_col``.

        ``shift_code`` is taken from row ``date_row + 3`` in the normal-hours
        column and is used to determine NS/NT when there is no OT.
        """

        current_row = start_row
        records_written = 0

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

            if _is_alphanumeric(work_val) and _has_positive_hours(hours_val) and str(work_val) != "0":
                # Columns A, B, C
                worksheet.write(current_row, 0, "A")
                worksheet.write(current_row, 1, "N")
                worksheet.write(current_row, 2, staff_id, staffid_right_format)

                # Column E: WorkOrderNum (with thin border)
                worksheet.write(current_row, 4, work_val, data_border_format)

                # Column G: TransDate
                worksheet.write(current_row, 6, date_value)

                # Column H: Manhour (nCol value). Non-empty cells should
                # have thin borders, be centred, and use Aptos font.
                if _has_positive_hours(hours_val):
                    worksheet.write(current_row, 7, float(hours_val), data_border_format)

                # Column I: TransType.
                # Priority order:
                #   1) If there is OT (value in ot_col), TransType = "OT".
                #   2) Otherwise, derive from shift code on the header row:
                #        N  -> NS
                #        not N -> NT
                # Column I cells have no borders.
                trans_type = ""

                # OT condition (highest precedence).
                if ot_val not in (None, "") and _has_positive_hours(ot_val):
                    trans_type = "OT"
                else:
                    # Fallback to shift-based NS/NT logic.
                    if shift_code not in (None, ""):
                        shift_str = str(shift_code).strip().upper()
                        trans_type = "NS" if shift_str == "N" else "NT"

                if trans_type:
                    worksheet.write(current_row, 8, trans_type)

                current_row += 1
                records_written += 1

            row += 1

        return records_written

    # For each staff/column-set combination we create a separate worksheet and
    # walk through the supplied list of candidate date rows. For each row we
    # process the date in column C using column set 1, then the date in column
    # J using column set 2. Both cells are assumed to contain dates.
    for idx, staff in enumerate(staffID):
        worksheet = workbook.add_worksheet(str(staff))

        # Set up headers and column widths for this worksheet.
        _initialise_worksheet(worksheet)

        current_row = data_start_row
        records_written = 0

        for date_row in dates:
            c_val = src_ws[f"C{date_row}"].value
            if c_val not in (None, ""):
                shift_code_c = src_ws[f"{nCol1[idx]}{date_row + 3}"].value
                written = _process_date_block(
                    worksheet,
                    staff,
                    current_row,
                    date_row,
                    c_val,
                    jCol1,
                    nCol1[idx],
                    oCol1[idx],
                    shift_code_c,
                )
                current_row += written
                records_written += written

            j_val = src_ws[f"J{date_row}"].value
            if j_val not in (None, ""):
                shift_code_j = src_ws[f"{nCol2[idx]}{date_row + 3}"].value
                written = _process_date_block(
                    worksheet,
                    staff,
                    current_row,
                    date_row,
                    j_val,
                    jCol2,
                    nCol2[idx],
                    oCol2[idx],
                    shift_code_j,
                )
                current_row += written
                records_written += written

        # Add an autofilter that covers the header row and all data rows on
        # this worksheet. Using zero-based (row, col) indices here.
        last_filter_row = data_start_row + records_written if records_written else 0
        worksheet.autofilter(0, 0, last_filter_row, 11)

    # Finalize and save the workbook.
    workbook.close()
    src_wb.close()


def main() -> None:
    """Entry point when running this script directly."""

    # Staff ID used for every record written to columns A, B, C.
    staffID = ["309605", "429465", "488704"]  # TODO: replace with the actual staff id as needed.

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
    nCol1, nCol2 = ["C", "E", "G"], ["J", "L", "N"]
    oCol1, oCol2 = ["D", "F", "H"], ["K", "M", "O"]
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
