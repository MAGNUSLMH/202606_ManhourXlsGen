from __future__ import annotations

from datetime import date, datetime
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

    if not (
        len(staffID) == len(nCol1) == len(nCol2) == len(oCol1) == len(oCol2)
    ):
        raise ValueError(
            "staffID, nCol1, nCol2, oCol1, and oCol2 must all have the same length"
        )

    # Open the source timesheet from which dates and working orders are read.
    src_wb = load_workbook(timesheet_path, data_only=True)
    try:
        src_ws = src_wb["Timesheet"]
    except KeyError:
        src_ws = src_wb.active

    # Create the workbook for the output.
    workbook = xlsxwriter.Workbook(filename)

    # Common formatting properties for all header cells.
    base_header_props = {
        "bold": True,
        "font_name": "Aptos",
        "underline": 1,
        "bg_color": "#D9D9D9",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }

    red_header_format = workbook.add_format({
        **base_header_props,
        "font_color": "#C00000",
    })

    black_header_format = workbook.add_format({
        **base_header_props,
        "font_color": "#000000",
    })

    date_format = workbook.add_format({
        "num_format": "m/d/yyyy",
        "align": "center",
    })

    headers = [
        ("Action", "red"),
        ("AdjInd", "red"),
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

    def _initialise_worksheet(ws) -> None:
        for col_idx, (label, color) in enumerate(headers):
            cell_format = (
                red_header_format if color == "red" else black_header_format
            )
            ws.write(0, col_idx, label, cell_format)

        ws.set_column("A:A", 9)
        ws.set_column("B:B", 9)
        ws.set_column("C:C", 10)
        ws.set_column("D:D", 12)
        ws.set_column("E:E", 14)
        ws.set_column("F:F", 11)
        ws.set_column("G:G", 11, date_format)
        ws.set_column("H:H", 10)
        ws.set_column("I:I", 11)
        ws.set_column("J:J", 5)
        ws.set_column("K:K", 9)
        ws.set_column("L:L", 13)

    data_border_format = workbook.add_format({
        "border": 1,
        "align": "center",
        "font_name": "Aptos",
    })

    staffid_right_format = workbook.add_format({"align": "right"})
    data_start_row = 1

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
        current_row = start_row
        records_written = 0
        cl_mode = False

        row = date_row + 7
        while row <= src_ws.max_row:
            work_val = src_ws[f"{work_col}{row}"].value
            hours_val = src_ws[f"{hours_col}{row}"].value
            ot_val = src_ws[f"{ot_col}{row}"].value

            marker_val = src_ws[f"A{row}"].value
            if isinstance(marker_val, str):
                marker_str = marker_val.strip()
                if (
                    "其他時間代號" in marker_str
                    or "other hours code" in marker_str.lower()
                ):
                    cl_mode = True

            if work_val is None and hours_val is None and ot_val is None:
                if row >= date_row + 52:
                    break
                row += 1
                continue

            if (
                _is_alphanumeric(work_val)
                and _has_positive_hours(hours_val)
                and str(work_val) != "0"
            ):
                worksheet.write(current_row, 0, "A")
                worksheet.write(current_row, 1, "N")
                worksheet.write(current_row, 2, staff_id, staffid_right_format)
                worksheet.write(current_row, 4, work_val, data_border_format)
                worksheet.write(current_row, 6, date_value)

                if _has_positive_hours(hours_val):
                    worksheet.write(
                        current_row, 7, float(hours_val), data_border_format
                    )

                trans_type = ""
                if ot_val not in (None, "") and _has_positive_hours(ot_val):
                    trans_type = "OT"
                else:
                    if shift_code not in (None, ""):
                        shift_str = str(shift_code).strip().upper()
                        if cl_mode:
                            trans_type = "CL"
                        else:
                            trans_type = "NS" if shift_str == "N" else "NT"

                if trans_type:
                    worksheet.write(current_row, 8, trans_type)

                current_row += 1
                records_written += 1

            row += 1

        return records_written

    for idx, staff in enumerate(staffID):
        worksheet = workbook.add_worksheet(str(staff))
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

        last_filter_row = (
            data_start_row + records_written if records_written else 0
        )
        worksheet.autofilter(0, 0, last_filter_row, 11)

    workbook.close()
    src_wb.close()


def main() -> None:
    """Entry point placeholder.

    Configs are now handled dynamically via Flask upload context.
    """
    pass


if __name__ == "__main__":
    main()