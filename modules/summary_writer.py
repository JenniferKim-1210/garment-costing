import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from io import BytesIO


def create_summary_excel(styles_cost_data: dict) -> BytesIO:
    """Create a clean summary costing Excel with all calculations done."""
    wb = openpyxl.Workbook()

    _write_summary_sheet(wb, styles_cost_data)
    _write_fabric_detail_sheet(wb, styles_cost_data)
    _write_trim_detail_sheet(wb, styles_cost_data)

    # Remove default empty sheet if we created named ones
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        del wb["Sheet"]

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def _write_summary_sheet(wb, styles_cost_data: dict):
    ws = wb.create_sheet("Cost Summary")

    header_font = Font(bold=True, size=10, name="Arial")
    data_font = Font(size=9, name="Arial")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font_white = Font(bold=True, size=10, name="Arial", color="FFFFFF")
    money_fmt = '#,##0.000'
    pct_fmt = '0.0%'

    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )

    headers = [
        "Style#", "Style Name", "CM/FTY", "SMV", "CM",
        "Fabric (Buyer)", "Fabric (Net)", "Fabric Transit", "Fabric Total (B)",
        "Trim (Buyer)", "Trim (Net)", "Trim Transit", "Trim Total (B)",
        "Graphic", "Wash", "Packing", "Misc",
        "FOB Buyer", "FOB Net", "Diff",
    ]

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    row = 2
    for style_num, data in styles_cost_data.items():
        def _w(col, val, fmt=None, bold=False):
            cell = ws.cell(row=row, column=col, value=val)
            cell.font = Font(bold=bold, size=9, name="Arial") if bold else data_font
            cell.border = thin_border
            if fmt:
                cell.number_format = fmt

        _w(1, style_num)
        _w(2, data.get("style_name", ""))
        _w(3, data.get("cm_fty", ""))
        _w(4, data.get("smv", 0))
        _w(5, data.get("cm", 0), money_fmt)
        _w(6, data.get("fabric_cost_buyer", data.get("fabric_cost", 0)), money_fmt)
        _w(7, data.get("fabric_cost_net", 0), money_fmt)
        _w(8, data.get("fabric_transit", 0), money_fmt)
        _w(9, data.get("fabric_total", 0), money_fmt)
        _w(10, data.get("trim_cost_buyer", data.get("trim_cost", 0)), money_fmt)
        _w(11, data.get("trim_cost_net", 0), money_fmt)
        _w(12, data.get("trim_transit", 0), money_fmt)
        _w(13, data.get("trim_total", 0), money_fmt)
        _w(14, data.get("graphic_price", 0), money_fmt)
        _w(15, data.get("wash_price", 0), money_fmt)
        _w(16, data.get("packing", 0), money_fmt)
        _w(17, data.get("misc", 0), money_fmt)
        _w(18, data.get("fob", 0), money_fmt, bold=True)
        _w(19, data.get("fob_net", 0), money_fmt, bold=True)
        _w(20, round(data.get("fob", 0) - data.get("fob_net", 0), 3), money_fmt)

        row += 1

    # Column widths
    widths = [12, 35, 10, 8, 10, 12, 12, 10, 12, 12, 12, 10, 12, 10, 10, 10, 10, 12, 12, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


def _write_fabric_detail_sheet(wb, styles_cost_data: dict):
    ws = wb.create_sheet("Fabric Detail")

    header_font = Font(bold=True, size=9, name="Arial", color="FFFFFF")
    data_font = Font(size=9, name="Arial")
    header_fill = PatternFill("solid", fgColor="548235")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    money_fmt = '#,##0.000'

    headers = [
        "Style#", "Placement", "Item", "Detail", "Color",
        "Mill", "C/O", "Shrinkage", "YY", "Loss%",
        "Size Upcharge%", "YY Final", "U/P (Buyer)", "Cost",
        "Tariff%", "After Tariff", "Transit", "Total",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border

    row = 2
    for style_num, data in styles_cost_data.items():
        for fb in data.get("fabric_details", []):
            ws.cell(row=row, column=1, value=style_num).font = data_font
            ws.cell(row=row, column=2, value=fb.get("placement", "")).font = data_font
            ws.cell(row=row, column=3, value=fb.get("item", "")).font = data_font
            ws.cell(row=row, column=4, value=fb.get("detail", "")).font = data_font
            ws.cell(row=row, column=5, value=fb.get("color_type", "")).font = data_font
            ws.cell(row=row, column=6, value=fb.get("mill", "")).font = data_font
            ws.cell(row=row, column=7, value=fb.get("c_o", "")).font = data_font
            ws.cell(row=row, column=8, value=fb.get("shrinkage", "")).font = data_font
            ws.cell(row=row, column=9, value=fb.get("yy", 0)).font = data_font
            ws.cell(row=row, column=10, value=fb.get("loss", 0)).font = data_font
            ws.cell(row=row, column=11, value=fb.get("size_upcharge", 0)).font = data_font
            ws.cell(row=row, column=12, value=fb.get("yy_final", 0)).font = data_font
            ws.cell(row=row, column=12).number_format = money_fmt
            ws.cell(row=row, column=13, value=fb.get("buyer_price", 0)).font = data_font
            ws.cell(row=row, column=13).number_format = money_fmt
            ws.cell(row=row, column=14, value=fb.get("cost", 0)).font = data_font
            ws.cell(row=row, column=14).number_format = money_fmt
            ws.cell(row=row, column=15, value=fb.get("tariff", 0)).font = data_font
            ws.cell(row=row, column=16, value=fb.get("after_tariff", 0)).font = data_font
            ws.cell(row=row, column=16).number_format = money_fmt
            ws.cell(row=row, column=17, value=fb.get("transit", 0)).font = data_font
            ws.cell(row=row, column=17).number_format = money_fmt
            ws.cell(row=row, column=18, value=fb.get("total", 0)).font = data_font
            ws.cell(row=row, column=18).number_format = money_fmt

            for c in range(1, 19):
                ws.cell(row=row, column=c).border = thin_border
            row += 1


def _write_trim_detail_sheet(wb, styles_cost_data: dict):
    ws = wb.create_sheet("Trim Detail")

    header_font = Font(bold=True, size=9, name="Arial", color="FFFFFF")
    data_font = Font(size=9, name="Arial")
    header_fill = PatternFill("solid", fgColor="BF8F00")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    money_fmt = '#,##0.000'

    headers = [
        "Style#", "Placement", "Item", "Detail", "Supplier",
        "C/O", "YY", "Loss%", "YY Final", "U/P (Buyer)",
        "Cost", "Tariff%", "After Tariff", "Transit", "Total",
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = thin_border

    row = 2
    for style_num, data in styles_cost_data.items():
        for tr in data.get("trim_details", []):
            ws.cell(row=row, column=1, value=style_num).font = data_font
            ws.cell(row=row, column=2, value=tr.get("placement", "")).font = data_font
            ws.cell(row=row, column=3, value=tr.get("item", "")).font = data_font
            ws.cell(row=row, column=4, value=tr.get("detail", "")).font = data_font
            ws.cell(row=row, column=5, value=tr.get("supplier", "")).font = data_font
            ws.cell(row=row, column=6, value=tr.get("c_o", "")).font = data_font
            ws.cell(row=row, column=7, value=tr.get("yy", 0)).font = data_font
            ws.cell(row=row, column=8, value=tr.get("loss", 0)).font = data_font
            ws.cell(row=row, column=9, value=tr.get("yy_final", 0)).font = data_font
            ws.cell(row=row, column=9).number_format = money_fmt
            ws.cell(row=row, column=10, value=tr.get("buyer_price", 0)).font = data_font
            ws.cell(row=row, column=10).number_format = money_fmt
            ws.cell(row=row, column=11, value=tr.get("cost", 0)).font = data_font
            ws.cell(row=row, column=11).number_format = money_fmt
            ws.cell(row=row, column=12, value=tr.get("tariff", 0)).font = data_font
            ws.cell(row=row, column=13, value=tr.get("after_tariff", 0)).font = data_font
            ws.cell(row=row, column=13).number_format = money_fmt
            ws.cell(row=row, column=14, value=tr.get("transit", 0)).font = data_font
            ws.cell(row=row, column=14).number_format = money_fmt
            ws.cell(row=row, column=15, value=tr.get("total", 0)).font = data_font
            ws.cell(row=row, column=15).number_format = money_fmt

            for c in range(1, 16):
                ws.cell(row=row, column=c).border = thin_border
            row += 1
