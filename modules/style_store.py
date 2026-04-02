"""Per-style version management.

Two storage backends:
1. JSON files in data/saved_styles/ (internal, fast)
2. Single Excel file on user-specified path (portable, visible)

Excel structure:
    Sheet "Index" - master list of all style versions
    Sheet "{style}_{version}" - fabric/trim/settings detail per version
"""

import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pathlib import Path
from datetime import datetime
import shutil

STYLES_DIR = Path(__file__).parent.parent / "data" / "saved_styles"
STYLES_DIR.mkdir(parents=True, exist_ok=True)

DEFAULT_EXCEL_PATH = Path(__file__).parent.parent / "data" / "PINK_Costing_DB.xlsx"


def save_style(style_num: int, version: str, comment: str,
               settings: dict, fabrics: list, trims: list) -> Path:
    """Save a single style version."""
    style_dir = STYLES_DIR / str(style_num)
    style_dir.mkdir(parents=True, exist_ok=True)

    safe_version = _safe_name(version)
    data = {
        "style_num": style_num,
        "version": version,
        "comment": comment,
        "saved_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "settings": settings,
        "fabrics": fabrics,
        "trims": trims,
    }

    path = style_dir / f"{safe_version}.json"
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    return path


def load_style(style_num: int, version: str) -> dict | None:
    """Load a specific style version."""
    safe_version = _safe_name(version)
    path = STYLES_DIR / str(style_num) / f"{safe_version}.json"
    if not path.exists():
        return None
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def list_style_versions(style_num: int) -> list[dict]:
    """List all saved versions for a style, newest first."""
    style_dir = STYLES_DIR / str(style_num)
    if not style_dir.exists():
        return []

    versions = []
    for p in style_dir.glob("*.json"):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            versions.append({
                "version": data.get("version", p.stem),
                "comment": data.get("comment", ""),
                "saved_at": data.get("saved_at", ""),
                "filename": p.stem,
            })
        except Exception:
            pass

    versions.sort(key=lambda x: x["saved_at"], reverse=True)
    return versions


def list_all_styles() -> dict:
    """List all styles that have at least one saved version.
    Returns {style_num: [versions]}
    """
    if not STYLES_DIR.exists():
        return {}
    result = {}
    for d in sorted(STYLES_DIR.iterdir()):
        if d.is_dir():
            try:
                sn = int(d.name)
            except ValueError:
                continue
            versions = list_style_versions(sn)
            if versions:
                result[sn] = versions
    return result


def delete_style_version(style_num: int, version: str):
    """Delete a specific version."""
    safe_version = _safe_name(version)
    path = STYLES_DIR / str(style_num) / f"{safe_version}.json"
    if path.exists():
        path.unlink()


def copy_style_version(style_num: int, from_version: str, to_version: str) -> Path:
    """Copy an existing version to a new version name."""
    data = load_style(style_num, from_version)
    if data is None:
        raise ValueError(f"Version '{from_version}' not found for style {style_num}")

    data["version"] = to_version
    data["comment"] = f"Copied from '{from_version}'. " + data.get("comment", "")
    data["saved_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    return save_style(
        style_num, to_version, data["comment"],
        data["settings"], data["fabrics"], data["trims"],
    )


def _safe_name(name: str) -> str:
    """Convert version name to safe filename."""
    safe = name.strip().replace(" ", "_").replace("/", "-").replace("\\", "-")
    safe = "".join(c for c in safe if c.isalnum() or c in "_-.")
    return safe or "unnamed"


# ============================================================
# Excel File Backend - single .xlsx with all style versions
# ============================================================

def _sheet_name(style_num: int, version: str) -> str:
    """Generate sheet name from style + version (max 31 chars for Excel)."""
    s = f"{style_num}_{_safe_name(version)}"
    return s[:31]


def save_style_to_excel(style_num: int, version: str, comment: str,
                        settings: dict, fabrics: list, trims: list,
                        excel_path: str | Path = None) -> Path:
    """Save/update a style version into the Excel DB file."""
    excel_path = Path(excel_path) if excel_path else DEFAULT_EXCEL_PATH
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    sheet = _sheet_name(style_num, version)

    # Load or create workbook
    if excel_path.exists():
        wb = openpyxl.load_workbook(excel_path)
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    # Remove old sheet if exists (update)
    if sheet in wb.sheetnames:
        del wb[sheet]

    ws = wb.create_sheet(sheet)

    # Styles
    hdr_font = Font(bold=True, size=9, name="Arial")
    hdr_fill = PatternFill("solid", fgColor="D9E1F2")
    data_font = Font(size=9, name="Arial")
    comment_fill = PatternFill("solid", fgColor="FFFFCC")
    thin = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

    # Row 1: Header info
    ws["A1"] = "Style#"
    ws["B1"] = style_num
    ws["C1"] = "Version"
    ws["D1"] = version
    ws["E1"] = "Saved"
    ws["F1"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    for c in range(1, 7):
        ws.cell(1, c).font = hdr_font

    # Row 2: Comment
    ws["A2"] = "Comment"
    ws["B2"] = comment
    ws["A2"].font = hdr_font
    ws["B2"].fill = comment_fill
    ws.merge_cells("B2:F2")

    # Row 3: Settings
    ws["A3"] = "CV#"
    ws["B3"] = settings.get("cv_num", "")
    ws["C3"] = "SMV"
    ws["D3"] = settings.get("smv", 0)
    ws["E3"] = "CM/FTY"
    ws["F3"] = settings.get("cm_fty", "")
    ws["G3"] = "EPM"
    ws["H3"] = settings.get("epm", 0)
    ws["I3"] = "Packing"
    ws["J3"] = settings.get("packing", 0)
    ws["K3"] = "Wash"
    ws["L3"] = settings.get("wash_price", 0)
    for c in range(1, 13):
        ws.cell(3, c).font = hdr_font if c % 2 == 1 else data_font

    # Row 5: Fabric header
    fb_headers = ["Placement", "Item", "Mill", "C/O", "Shrinkage",
                  "YY", "Loss%", "Size Up%", "U/P Buyer", "U/P Net", "Tariff%"]
    ws["A5"] = "FABRIC"
    ws["A5"].font = Font(bold=True, size=10, name="Arial", color="FFFFFF")
    ws["A5"].fill = PatternFill("solid", fgColor="548235")
    for i, h in enumerate(fb_headers, 1):
        cell = ws.cell(6, i, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = thin

    # Fabric data
    for ri, fb in enumerate(fabrics, 7):
        for ci, key in enumerate(fb_headers, 1):
            val = fb.get(key, fb.get(key.lower().replace(" ", "_").replace("%", ""), ""))
            cell = ws.cell(ri, ci, val)
            cell.font = data_font
            cell.border = thin

    # Trim section
    trim_start = 7 + len(fabrics) + 2
    tr_headers = ["Placement", "Item", "Detail", "Supplier", "C/O",
                  "YY", "Loss%", "U/P Buyer", "U/P Net", "Tariff%", "RM Air"]
    ws.cell(trim_start - 1, 1, "TRIM & LABEL").font = Font(bold=True, size=10, name="Arial", color="FFFFFF")
    ws.cell(trim_start - 1, 1).fill = PatternFill("solid", fgColor="BF8F00")
    for i, h in enumerate(tr_headers, 1):
        cell = ws.cell(trim_start, i, h)
        cell.font = hdr_font
        cell.fill = hdr_fill
        cell.border = thin

    for ri, tr in enumerate(trims, trim_start + 1):
        for ci, key in enumerate(tr_headers, 1):
            val = tr.get(key, tr.get(key.lower().replace(" ", "_").replace("%", ""), ""))
            cell = ws.cell(ri, ci, val)
            cell.font = data_font
            cell.border = thin

    # Column widths
    widths = [20, 18, 12, 8, 10, 8, 8, 10, 10, 10, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Update/create Index sheet
    _update_index_sheet(wb, style_num, version, comment, settings)

    wb.save(excel_path)
    return excel_path


def _update_index_sheet(wb, style_num, version, comment, settings):
    """Maintain an Index sheet with all style versions summary."""
    if "Index" not in wb.sheetnames:
        ws = wb.create_sheet("Index", 0)
        headers = ["Style#", "Version", "Comment", "SMV", "CM/FTY", "Saved At", "Sheet Name"]
        hdr_font = Font(bold=True, size=10, name="Arial", color="FFFFFF")
        hdr_fill = PatternFill("solid", fgColor="4472C4")
        for i, h in enumerate(headers, 1):
            cell = ws.cell(1, i, h)
            cell.font = hdr_font
            cell.fill = hdr_fill
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 40
        ws.column_dimensions["D"].width = 8
        ws.column_dimensions["E"].width = 12
        ws.column_dimensions["F"].width = 18
        ws.column_dimensions["G"].width = 25
    else:
        ws = wb["Index"]

    sheet = _sheet_name(style_num, version)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    data_font = Font(size=9, name="Arial")

    # Check if this style+version already exists, update it
    found = False
    for row in range(2, ws.max_row + 1):
        if ws.cell(row, 7).value == sheet:
            ws.cell(row, 3, comment).font = data_font
            ws.cell(row, 4, settings.get("smv", 0)).font = data_font
            ws.cell(row, 5, settings.get("cm_fty", "")).font = data_font
            ws.cell(row, 6, now).font = data_font
            found = True
            break

    if not found:
        r = ws.max_row + 1
        ws.cell(r, 1, style_num).font = data_font
        ws.cell(r, 2, version).font = data_font
        ws.cell(r, 3, comment).font = data_font
        ws.cell(r, 4, settings.get("smv", 0)).font = data_font
        ws.cell(r, 5, settings.get("cm_fty", "")).font = data_font
        ws.cell(r, 6, now).font = data_font
        ws.cell(r, 7, sheet).font = data_font


def load_style_from_excel(style_num: int, version: str,
                          excel_path: str | Path = None) -> dict | None:
    """Load a style version from the Excel DB file."""
    excel_path = Path(excel_path) if excel_path else DEFAULT_EXCEL_PATH
    if not excel_path.exists():
        return None

    sheet = _sheet_name(style_num, version)
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if sheet not in wb.sheetnames:
        return None

    ws = wb[sheet]
    comment = ws["B2"].value or ""

    settings = {
        "cv_num": ws["B3"].value or "",
        "smv": ws["D3"].value or 0,
        "cm_fty": ws["F3"].value or "",
        "epm": ws["H3"].value or 0.11,
        "packing": ws["J3"].value or 0.12,
        "wash_price": ws["L3"].value or 0,
    }

    # Read fabric rows (start at row 7 until empty)
    fb_headers = [ws.cell(6, c).value for c in range(1, 12)]
    fabrics = []
    for r in range(7, ws.max_row + 1):
        vals = [ws.cell(r, c).value for c in range(1, 12)]
        if not any(vals):
            break
        fabrics.append(dict(zip(fb_headers, vals)))

    # Find trim section
    trims = []
    for r in range(7 + len(fabrics), ws.max_row + 1):
        if ws.cell(r, 1).value == "TRIM & LABEL":
            tr_headers = [ws.cell(r + 1, c).value for c in range(1, 12)]
            for tr_row in range(r + 2, ws.max_row + 1):
                vals = [ws.cell(tr_row, c).value for c in range(1, 12)]
                if not any(vals):
                    break
                trims.append(dict(zip(tr_headers, vals)))
            break

    return {
        "style_num": style_num,
        "version": version,
        "comment": comment,
        "settings": settings,
        "fabrics": fabrics,
        "trims": trims,
    }


def list_versions_from_excel(excel_path: str | Path = None) -> list[dict]:
    """List all style versions from the Excel Index sheet."""
    excel_path = Path(excel_path) if excel_path else DEFAULT_EXCEL_PATH
    if not excel_path.exists():
        return []

    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "Index" not in wb.sheetnames:
        return []

    ws = wb["Index"]
    results = []
    for r in range(2, ws.max_row + 1):
        sn = ws.cell(r, 1).value
        if sn is None:
            continue
        results.append({
            "style_num": int(sn),
            "version": ws.cell(r, 2).value or "",
            "comment": ws.cell(r, 3).value or "",
            "smv": ws.cell(r, 4).value or 0,
            "cm_fty": ws.cell(r, 5).value or "",
            "saved_at": ws.cell(r, 6).value or "",
            "sheet": ws.cell(r, 7).value or "",
        })
    return results


def list_style_versions_from_excel(style_num: int, excel_path: str | Path = None) -> list[dict]:
    """List all versions for a specific style from Excel."""
    all_versions = list_versions_from_excel(excel_path)
    return [v for v in all_versions if v["style_num"] == style_num]
