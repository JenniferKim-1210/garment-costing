import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from io import BytesIO
import copy


# PPK column mapping (1-indexed)
COL = {
    # Header row
    "style_num": 1,       # A
    "style_name": 2,      # B
    "color_code": 6,      # F
    "order_qty": 8,       # H
    "fob": 9,             # I
    "fob_after_tariff": 10,  # J
    "smv": 11,            # K
    "epm": 12,            # L
    "cm": 13,             # M
    "margin": 14,         # N
    # Fabric section
    "fb_placement": 15,   # O
    "fb_item": 16,        # P
    "fb_detail": 17,      # Q
    "fb_color": 18,       # R
    "fb_mill": 19,        # S
    "fb_co": 20,          # T
    "fb_shrinkage": 21,   # U
    "fb_yy": 22,          # V
    "fb_loss": 23,        # W
    "fb_size_upcharge": 24,  # X
    "fb_yy_final": 25,    # Y
    "fb_up": 26,          # Z
    "fb_cost": 27,        # AA
    "fb_tariff": 28,      # AB
    "fb_after_tariff": 29,  # AC
    "fb_rm_shipmode": 30, # AD
    "fb_transit_per_yd": 31,  # AE
    "fb_transit_cost": 32,  # AF
    "fb_total": 33,       # AG
    "fb_sum_cost": 34,    # AH
    "fb_sum_after_tariff": 35,  # AI
    # Trim section
    "tr_placement": 36,   # AJ
    "tr_item": 37,        # AK
    "tr_detail": 38,      # AL
    "tr_supplier": 39,    # AM
    "tr_co": 40,          # AN
    "tr_yy": 41,          # AO
    "tr_loss": 42,        # AP
    "tr_yy_final": 43,    # AQ
    "tr_up": 44,          # AR
    "tr_zipper_add": 45,  # AS
    "tr_zipper_upcharge": 46,  # AT
    "tr_cost": 47,        # AU
    "tr_tariff": 48,      # AV
    "tr_after_tariff": 49,  # AW
    "tr_rm_shipmode": 50, # AX
    "tr_transit_per_unit": 51,  # AY
    "tr_transit_cost": 52,  # AZ
    "tr_total": 53,       # BA
    "tr_sum_cost": 54,    # BB
    "tr_sum_after_tariff": 55,  # BC
    # Others
    "graphic_tech": 56,   # BD
    "graphic_price": 57,  # BE
    "wash_tech": 58,      # BF
    "wash_price": 59,     # BG
    "packing": 60,        # BH
    "misc": 61,           # BI
    "sales_margin": 62,   # BJ
    "business_margin": 63,  # BK
    "order_amount": 64,   # BL
}

BLOCK_SIZE = 17  # rows per style block (header rows + data rows + remarks)
FABRIC_ROWS = 8  # max fabric rows per block
TRIM_ROWS = 8    # max trim rows per block


def write_ppk_to_template(template_file, styles_cost_data: dict, start_row: int = 5) -> BytesIO:
    """Write cost data into existing PPK template.

    Args:
        template_file: file-like object of PPK template
        styles_cost_data: dict of {style_num: cost_summary_dict}
        start_row: first data row in PPK sheet

    Returns:
        BytesIO with the modified Excel file
    """
    wb = openpyxl.load_workbook(template_file)

    # Find the PPK sheet (try common names)
    ppk_sheet = None
    for name in wb.sheetnames:
        if "PPK" in name.upper():
            ppk_sheet = wb[name]
            break
    if ppk_sheet is None:
        ppk_sheet = wb.active

    # Find existing style blocks
    style_rows = _find_style_blocks(ppk_sheet, start_row)

    for style_num, cost_data in styles_cost_data.items():
        if style_num in style_rows:
            block_start = style_rows[style_num]
            _write_style_block(ppk_sheet, block_start, cost_data)

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def create_new_ppk(styles_cost_data: dict) -> BytesIO:
    """Create a fresh PPK Excel from scratch with all cost data."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PPK COST"

    # Header styles
    header_font = Font(bold=True, size=9, name="Arial")
    data_font = Font(size=9, name="Arial")
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    # Write headers (row 4)
    headers = {
        1: "S/#", 2: "Style Name", 6: "COLOR CODE",
        8: "ORDER\nQ'TY", 9: "FOB", 10: "FOB\nAFTER TARIFF",
        11: "SMV", 12: "EPM", 13: "CM", 14: "MARGIN",
        15: "PLACEMENT", 16: "FABRIC\nITEM", 17: "FABRIC DETAIL",
        18: "COLOR", 19: "MILL", 20: "C/O",
        21: "SHRINKAGE", 22: "FABRIC\nYY", 23: "FABRIC\nLOSS",
        24: "SIZE\nUPCHARGE", 25: "YY", 26: "FABRIC\nU/P",
        27: "FABRIC\nCOST", 28: "TARIFF\n%", 29: "AFTER\nTARIFF",
        30: "RM\nSHIPMODE", 31: "TRANSIT\nCOST/YD", 32: "RM TRANSIT\nCOST",
        33: "FABRIC+\nTRANSIT", 34: "FABRIC+\nTRANSIT", 35: "AFTER\nTARIFF",
        36: "PLACEMENT", 37: "ITEM", 38: "ACC\nITEM",
        39: "SUPPLIER", 40: "C/O", 41: "ACC\nYY",
        42: "ACC\nLOSS", 43: "YY", 44: "ACC\nU/P",
        45: "ZIPPER ADD", 46: "ZIPPER\nUPCHARGE",
        47: "ACC\nCOST", 48: "TARIFF\n%", 49: "AFTER\nTARIFF",
        50: "RM\nSHIPMODE", 51: "TRANSIT\nCOST", 52: "RM\nTRANSIT",
        53: "TRIM COST", 54: "ACC\nPRICE", 55: "AFTER TARIFF\n+TRANSIT",
        56: "GRAPHIC\nTECH", 57: "GRAPHIC\nPRICE",
        58: "WASH\nTECH", 59: "WASH\nPRICE",
        60: "PACKING", 61: "Misc",
    }

    for col, text in headers.items():
        cell = ws.cell(row=4, column=col, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Write style blocks
    current_row = 5
    for style_num, cost_data in styles_cost_data.items():
        current_row = _write_new_style_block(ws, current_row, style_num, cost_data, data_font, thin_border)
        current_row += 2  # gap between styles

    # Auto-fit column widths (approximate)
    for col in range(1, 62):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 12

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _find_style_blocks(ws, start_row: int) -> dict:
    """Find style numbers and their starting rows in PPK sheet."""
    style_rows = {}
    for row in range(start_row, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val is not None:
            try:
                style_num = int(val)
                if style_num > 10000000:  # valid style number
                    style_rows[style_num] = row
            except (ValueError, TypeError):
                pass
    return style_rows


def _write_style_block(ws, block_start: int, cost_data: dict):
    """Write cost data into an existing style block."""
    fabric_details = cost_data.get("fabric_details", [])
    trim_details = cost_data.get("trim_details", [])

    # Data rows start at block_start + 1 (block_start is style header)
    data_start = block_start + 1

    # Write fabric data
    for i, fb in enumerate(fabric_details[:FABRIC_ROWS]):
        row = data_start + i
        _set_if_empty(ws, row, COL["fb_placement"], fb.get("placement", ""))
        _set_if_empty(ws, row, COL["fb_item"], fb.get("item", ""))
        _set_if_empty(ws, row, COL["fb_detail"], fb.get("detail", ""))
        _set_if_empty(ws, row, COL["fb_color"], fb.get("color_type", ""))
        _set_if_empty(ws, row, COL["fb_mill"], fb.get("mill", ""))
        _set_if_empty(ws, row, COL["fb_co"], fb.get("c_o", ""))
        ws.cell(row=row, column=COL["fb_shrinkage"], value=fb.get("shrinkage", ""))
        ws.cell(row=row, column=COL["fb_yy"], value=fb.get("yy", 0))
        ws.cell(row=row, column=COL["fb_loss"], value=fb.get("loss", 0))
        ws.cell(row=row, column=COL["fb_size_upcharge"], value=fb.get("size_upcharge", 0))

    # Write trim data
    for i, tr in enumerate(trim_details[:TRIM_ROWS]):
        row = data_start + i
        _set_if_empty(ws, row, COL["tr_placement"], tr.get("placement", ""))
        _set_if_empty(ws, row, COL["tr_item"], tr.get("item", ""))
        _set_if_empty(ws, row, COL["tr_detail"], tr.get("detail", ""))
        _set_if_empty(ws, row, COL["tr_supplier"], tr.get("supplier", ""))
        _set_if_empty(ws, row, COL["tr_co"], tr.get("c_o", ""))
        ws.cell(row=row, column=COL["tr_yy"], value=tr.get("yy", 0))
        ws.cell(row=row, column=COL["tr_loss"], value=tr.get("loss", 0))

    # Write summary values
    ws.cell(row=block_start, column=COL["smv"], value=cost_data.get("smv", 0))
    ws.cell(row=data_start, column=COL["graphic_tech"], value=cost_data.get("graphic_tech", "N/A"))
    ws.cell(row=data_start, column=COL["wash_tech"], value=cost_data.get("wash_tech", "N/A"))
    ws.cell(row=data_start, column=COL["packing"], value=cost_data.get("packing", 0))


def _write_new_style_block(ws, start_row: int, style_num: int, cost_data: dict,
                           font, border) -> int:
    """Write a complete style block for new PPK. Returns next available row."""
    fabric_details = cost_data.get("fabric_details", [])
    trim_details = cost_data.get("trim_details", [])
    max_rows = max(len(fabric_details), len(trim_details), 2)

    # Style header row
    ws.cell(row=start_row, column=1, value=style_num).font = Font(bold=True, size=9, name="Arial")
    ws.cell(row=start_row, column=2, value=cost_data.get("style_name", ""))
    ws.cell(row=start_row, column=9, value=cost_data.get("fob", 0))
    ws.cell(row=start_row, column=11, value=cost_data.get("smv", 0))
    ws.cell(row=start_row, column=13, value=cost_data.get("cm", 0))

    # Data rows
    data_start = start_row + 1
    for i in range(max_rows):
        row = data_start + i

        # Fabric
        if i < len(fabric_details):
            fb = fabric_details[i]
            ws.cell(row=row, column=COL["fb_placement"], value=fb.get("placement", ""))
            ws.cell(row=row, column=COL["fb_item"], value=fb.get("item", ""))
            ws.cell(row=row, column=COL["fb_detail"], value=fb.get("detail", ""))
            ws.cell(row=row, column=COL["fb_color"], value=fb.get("color_type", ""))
            ws.cell(row=row, column=COL["fb_mill"], value=fb.get("mill", ""))
            ws.cell(row=row, column=COL["fb_co"], value=fb.get("c_o", ""))
            ws.cell(row=row, column=COL["fb_shrinkage"], value=fb.get("shrinkage", ""))
            ws.cell(row=row, column=COL["fb_yy"], value=fb.get("yy", 0))
            ws.cell(row=row, column=COL["fb_loss"], value=fb.get("loss", 0))
            ws.cell(row=row, column=COL["fb_size_upcharge"], value=fb.get("size_upcharge", 0))
            ws.cell(row=row, column=COL["fb_yy_final"], value=fb.get("yy_final", 0))
            ws.cell(row=row, column=COL["fb_up"], value=fb.get("buyer_price", 0))
            ws.cell(row=row, column=COL["fb_cost"], value=fb.get("cost", 0))
            ws.cell(row=row, column=COL["fb_tariff"], value=fb.get("tariff", 0))
            ws.cell(row=row, column=COL["fb_after_tariff"], value=fb.get("after_tariff", 0))
            ws.cell(row=row, column=COL["fb_transit_cost"], value=fb.get("transit", 0))
            ws.cell(row=row, column=COL["fb_total"], value=fb.get("total", 0))

        # Trim
        if i < len(trim_details):
            tr = trim_details[i]
            ws.cell(row=row, column=COL["tr_placement"], value=tr.get("placement", ""))
            ws.cell(row=row, column=COL["tr_item"], value=tr.get("item", ""))
            ws.cell(row=row, column=COL["tr_detail"], value=tr.get("detail", ""))
            ws.cell(row=row, column=COL["tr_supplier"], value=tr.get("supplier", ""))
            ws.cell(row=row, column=COL["tr_co"], value=tr.get("c_o", ""))
            ws.cell(row=row, column=COL["tr_yy"], value=tr.get("yy", 0))
            ws.cell(row=row, column=COL["tr_loss"], value=tr.get("loss", 0))
            ws.cell(row=row, column=COL["tr_yy_final"], value=tr.get("yy_final", 0))
            ws.cell(row=row, column=COL["tr_up"], value=tr.get("buyer_price", 0))
            ws.cell(row=row, column=COL["tr_cost"], value=tr.get("cost", 0))
            ws.cell(row=row, column=COL["tr_tariff"], value=tr.get("tariff", 0))
            ws.cell(row=row, column=COL["tr_after_tariff"], value=tr.get("after_tariff", 0))
            ws.cell(row=row, column=COL["tr_transit_cost"], value=tr.get("transit", 0))
            ws.cell(row=row, column=COL["tr_total"], value=tr.get("total", 0))

    # Summary row (CM/FTY row)
    summary_row = data_start + max_rows
    ws.cell(row=summary_row, column=1, value="CM/FTY")
    ws.cell(row=summary_row, column=13, value=cost_data.get("cm", 0))
    ws.cell(row=summary_row, column=COL["fb_sum_cost"], value=cost_data.get("fabric_total", 0))
    ws.cell(row=summary_row, column=COL["tr_sum_cost"], value=cost_data.get("trim_total", 0))
    ws.cell(row=summary_row, column=COL["graphic_price"], value=cost_data.get("graphic_price", 0))
    ws.cell(row=summary_row, column=COL["wash_price"], value=cost_data.get("wash_price", 0))
    ws.cell(row=summary_row, column=COL["packing"], value=cost_data.get("packing", 0))

    # Apply font to all cells
    for r in range(start_row, summary_row + 1):
        for c in range(1, 62):
            cell = ws.cell(row=r, column=c)
            if cell.font == Font():
                cell.font = font

    return summary_row + 1


def _set_if_empty(ws, row, col, value):
    """Only set cell value if it's currently empty."""
    cell = ws.cell(row=row, column=col)
    if cell.value is None or str(cell.value).strip() == "":
        cell.value = value
