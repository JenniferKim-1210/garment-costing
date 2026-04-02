import pandas as pd
from pathlib import Path


def parse_bom(file) -> dict:
    """Parse BOM Extract Excel and return structured data per style."""
    sheets = pd.read_excel(file, sheet_name=None)

    bom_summary = sheets.get("BOM Summary")
    graphic_sheet = sheets.get("Graphic")
    color_detail = sheets.get("Fabric,Trim,Label for all color")
    all_sheet = sheets.get("All")

    if bom_summary is None:
        raise ValueError("BOM Summary sheet not found in the uploaded file.")

    styles = {}
    for style_num, group in bom_summary.groupby("Style#"):
        style_num = int(style_num)

        fabrics = group[group["Section"] == "Fabric"].reset_index(drop=True)
        trims = group[group["Section"] == "Trim"].reset_index(drop=True)
        labels = group[group["Section"] == "Labeling"].reset_index(drop=True)

        fabric_list = []
        for _, row in fabrics.iterrows():
            fabric_list.append({
                "placement": _clean(row.get("Placement", "")),
                "item": _clean(row.get("Item", "")),
                "material_code": _clean(row.get("Material Code", "")),
                "composition": _clean(row.get("Composition", "")),
                "detail": _clean(row.get("Detail", "")),
                "part_type": _clean(row.get("Part Type", "")),
                "yy": 0.0,
                "loss": 0.03,
                "shrinkage": "",
                "size_upcharge": 0.0,
            })

        trim_list = []
        for _, row in trims.iterrows():
            trim_list.append({
                "placement": _clean(row.get("Placement", "")),
                "item": _clean(row.get("Item", "")),
                "material_code": _clean(row.get("Material Code", "")),
                "detail": _clean(row.get("Detail", "")),
                "part_type": _clean(row.get("Part Type", "")),
                "yy": 1.0,
                "loss": 0.03,
            })

        label_list = []
        for _, row in labels.iterrows():
            label_list.append({
                "placement": _clean(row.get("Placement", "")),
                "item": _clean(row.get("Item", "")),
                "material_code": _clean(row.get("Material Code", "")),
                "detail": _clean(row.get("Detail", "")),
                "part_type": _clean(row.get("Part Type", "")),
                "yy": 1.0,
                "loss": 0.03,
            })

        graphics = []
        if graphic_sheet is not None:
            style_graphics = graphic_sheet[graphic_sheet["Style#"] == style_num]
            for _, row in style_graphics.iterrows():
                graphics.append({
                    "colorway": _clean(row.get("Colorway", "")),
                    "color_code": _clean(row.get("Color Code", "")),
                    "graphic_code": _clean(row.get("Graphic Code", "")),
                    "graphic_detail": _clean(row.get("Graphic Full Detail", "")),
                })

        colorways = []
        if color_detail is not None:
            style_colors = color_detail[color_detail["Style#"] == style_num]
            for _, row in style_colors.iterrows():
                colorways.append({
                    "colorway": _clean(row.get("Colorway", "")),
                    "cw_code": _clean(row.get("CW Code", "")),
                    "section": _clean(row.get("Section", "")),
                    "part_type": _clean(row.get("Part Type", "")),
                    "placement": _clean(row.get("Placement", "")),
                    "material_code": _clean(row.get("Material Code", "")),
                    "color_value": _clean(row.get("Color Value", "")),
                    "composition": _clean(row.get("Composition", "")),
                    "detail": _clean(row.get("Detail", "")),
                })

        styles[style_num] = {
            "fabrics": fabric_list,
            "trims": trim_list,
            "labels": label_list,
            "graphics": graphics,
            "colorways": colorways,
        }

    return styles


def get_style_names(file) -> dict:
    """Extract style number -> style name mapping from BOM."""
    sheets = pd.read_excel(file, sheet_name=None)
    all_sheet = sheets.get("All")
    if all_sheet is not None and "Style#" in all_sheet.columns:
        # Try to get names from final bom or all sheet
        pass

    bom_summary = sheets.get("BOM Summary")
    if bom_summary is not None:
        style_names = {}
        for style_num in bom_summary["Style#"].unique():
            style_names[int(style_num)] = str(style_num)
        return style_names
    return {}


def _clean(val) -> str:
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip()
