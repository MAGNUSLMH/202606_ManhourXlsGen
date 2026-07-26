from __future__ import annotations

from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side


def _is_alphanumeric(value) -> bool:
    """Return True if value is a non-empty string of only A–Z / a–z / 0–9.

    Used to detect working order codes in the timesheet columns.
    """

    if value is None:
        return False
    s = str(value).strip()
    return bool(s) and s.isalnum()


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


def highlight_workbook(
    filename: str = "sample.xlsx",
    staffID: list[str] | None = None,
    dates: list[str] | None = None,
    *,
    timesheet_path: str,
    jCol1: str,
    jCol2: str,
    nCol1: list[str],
    nCol2: list[str],
    oCol1: list[str],
    oCol2: list[str],
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

    sheet_names = src_wb.sheetnames[::]
    sheet_names.remove("Timesheet")
    for n in sheet_names:
        src_wb.remove(src_wb[n])

    # Create the workbook for the output. Individual worksheets will be
    # created per staff / column-set combination.
    # workbook = src_wb

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

    # Define two formats: one with orange font and one with green font.
    orange_font = Font(name="Aptos", size=24, color="FFFEB264")
    orange_fill = PatternFill(
        start_color="FFFFC000",  # Hex color FFC000 with FF prefix for opacity
        end_color="FFFFC000", 
        fill_type="solid"
    )
    green_font = Font(name="Aptos", size=24, color="FF85CC7E")
    black_font = Font(name="Aptos", size=24, bold=True, underline="single", color="FF000000")
    header_alignment = Alignment(horizontal="left", vertical="center")

    thin_side = Side(border_style="thin", color="FF000000")
    header_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    headers = [
        ("Date", black_font),
        ("", black_font),
        ("Remark", black_font)
    ]

    # Helper to set up a worksheet with header row and column widths.
    def _initialise_worksheet(ws) -> None:
        # Write the header labels to row 1 (index 0) with appropriate formatting.
        # This recreates the styled header row shown in the screenshot.
        # Assuming 'ws' is your active openpyxl worksheet object
        row_num = 1

        for col_idx, (header_name, font_style) in enumerate(headers, start=1):
            cell = ws.cell(row=row_num, column=col_idx)
            
            # Set the value
            cell.value = header_name
            
            # Apply individual style components
            cell.font = font_style
            cell.alignment = header_alignment
            cell.border = header_border

        # Optionally adjust column widths so that the text is visible and similar
        # to the appearance in the screenshot. This is not strictly required
        # for functionality but makes the result easier to inspect.
        ws.column_dimensions['A'].width = 40   # Date
        ws.column_dimensions['B'].width = 5   # Empty
        ws.column_dimensions['C'].width = 100  # Remark
        ws.auto_filter.ref = "A1:C100"


    data_start_row = 2

    def _process_date_block(
        staff_id: str,
        start_row: int,
        date_row: int,
        date_value,
        work_col: str,
        hours_col: str,
        ot_col: str,
    ) -> int:
        """Scan downward from a given date row for working orders in *work_col*.

        For every qualifying working order row we emit one record into the
        output sheet, using ``date_value`` as the TransDate and the matching
        hours / OT cells from ``hours_col`` and ``ot_col``.

        ``shift_code`` is taken from row ``date_row + 3`` in the normal-hours
        column and is used to determine NS/NT when there is no OT.
        """

        dct = {
            "reg_right":False,
            "other_left":False,
            "other_right":False,
            "ot_claim":False
        }
        dct_chinese = {
            "reg_right":"恆常工作編號 右➡️➡️手邊有數字",
            "other_left":"其他時間代號 左⬅️⬅️手邊有數字",
            "other_right":"其他時間代號 右➡️➡️手邊有數字",
            "ot_claim":"OT Claim 有數字"
        }
        current_row = start_row
        records_written = 0
        ot_val = src_ws[f"{ot_col}{date_row + 49}"].value
        comp_leave_code = "CL" if ot_val == "C/L" else ot_val

        # "CL" override state for this date block. For each day (date row)
        # we scan rows in the jCol area; if we encounter a marker row in
        # column A containing "其他時間代號" / "Other Hours Code", all
        # *subsequent* records emitted by this function for that day will
        # use TransType = "CL". This flag is local to the function, so it
        # resets automatically when a new day is processed.
        cl_mode = False

        row = date_row + 7
        while row <= src_ws.max_row:
            work_val = src_ws[f"{work_col}{row}"].value
            hours_val = src_ws[f"{hours_col}{row}"].value
            ot_val = src_ws[f"{ot_col}{row}"].value

            # Detect the "Other Hours Code" marker in column A. Once this
            # is seen for the current date block, all *subsequent* records
            # get TransType = "CL" regardless of OT or shift code.
            marker_val = src_ws[f"A{row}"].value
            

            if isinstance(marker_val, str):
                marker_str = marker_val.strip()
                if "其他時間代號" in marker_str: # or "other hours code" in marker_str.lower():
                    cl_mode = True
                elif ("Overtime Claimed" in marker_str or "granted" in marker_str.lower()) and _has_positive_hours(ot_val):
                    # ot claim: highlight if non-zero
                    src_ws[f"{ot_col}{row}"].fill = orange_fill
                    dct["ot_claim"] = True

            # If an entire trio of cells is empty we allow for a limited number
            # of blank rows; after that we stop scanning this block.
            if work_val is None and hours_val is None and ot_val is None:
                if row >= date_row + 52:
                    break
                row += 1
                continue

            if _is_alphanumeric(work_val) and _has_positive_hours(hours_val) and str(work_val) != "0":
                if cl_mode:
                    # cl_mode: fill left
                    src_ws[f"{hours_col}{row}"].fill = orange_fill
                    dct["other_left"] = True
                if _has_positive_hours(ot_val):
                    # ot: highlight in whichever mode
                    src_ws[f"{ot_col}{row}"].fill = orange_fill
                    if cl_mode:
                        dct["other_right"] = True
                    else:
                        dct["reg_right"] = True
            row += 1
        
        if list(dct.values()).count(True) == 0:
            write_ws = src_wb[staff_id]
            tmp_cell = write_ws.cell(row=current_row, column=1)
            tmp_cell.value = date_value
            tmp_cell.font = green_font
            tmp_cell.number_format = "m/d/yyyy"
            tmp_cell.alignment = Alignment(horizontal="left")
            
            current_row += 1
            records_written += 1
        else:
            write_ws = src_wb[staff_id]
            tmp_cell = write_ws.cell(row=current_row, column=1)
            tmp_cell.value = date_value
            tmp_cell.font = orange_font
            tmp_cell.number_format = "m/d/yyyy"
            tmp_cell.alignment = Alignment(horizontal="left")

            write_ws = src_wb[staff_id]
            tmp_cell = write_ws.cell(row=current_row, column=3)
            tmp_cell.value = "\n".join(dct_chinese[k] for k, v in dct.items() if v)
            tmp_cell.font = orange_font
            tmp_cell.alignment = Alignment(horizontal="left", wrap_text=True)
            current_row += 1
            records_written += 1

        return records_written

    # For each staff/column-set combination we create a separate worksheet and
    # walk through the supplied list of candidate date rows. For each row we
    # process the date in column C using column set 1, then the date in column
    # J using column set 2. Both cells are assumed to contain dates.
    for idx, staff in enumerate(staffID):
        worksheet = src_wb.create_sheet(title=staff)

        # Set up headers and column widths for this worksheet.
        _initialise_worksheet(worksheet)

        current_row = data_start_row
        records_written = 0

        for date_row in dates:
            c_val = src_ws[f"C{date_row}"].value
            if c_val not in (None, ""):
                written = _process_date_block(
                    staff,
                    current_row,
                    date_row,
                    c_val,
                    jCol1,
                    nCol1[idx],
                    oCol1[idx],
                )
                current_row += written
                records_written += written

            j_val = src_ws[f"J{date_row}"].value
            if j_val not in (None, ""):
                written = _process_date_block(
                    staff,
                    current_row,
                    date_row,
                    j_val,
                    jCol2,
                    nCol2[idx],
                    oCol2[idx],
                )
                current_row += written
                records_written += written

    # Finalize and save the workbook.
    src_wb.save(timesheet_path)
    src_wb.close()


def main() -> None:
    """Entry point when running this script directly."""

    # Staff ID used for every record written to columns A, B, C.
    staffID = ["488704", "429465", "676139"]  # TODO: replace with the actual staff id as needed.

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
    timesheet_path = "fjord.xlsx"

    # Create the output workbook using the configured parameters.
    highlight_workbook(
        datetime.now().strftime("%Y%m%d%H%M%S")+"_Highlight.xlsx",
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
