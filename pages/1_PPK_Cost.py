import streamlit as st
import pandas as pd
import json
from pathlib import Path
from io import BytesIO
from datetime import datetime
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from modules.bom_parser import parse_bom
from modules.price_matcher import (
    load_fabric_prices, load_trim_prices,
    match_fabric_price, match_trim_price,
)
from modules.cost_calculator import (
    calc_fabric_yy_final, calc_trim_yy_final,
    calc_after_tariff, calc_transit_cost, calc_fob, calc_misc,
)
from modules.ppk_writer import create_new_ppk, write_ppk_to_template
from modules.summary_writer import create_summary_excel
from modules.style_store import (
    save_style, load_style, list_style_versions,
    save_style_to_excel, load_style_from_excel,
    list_style_versions_from_excel, list_versions_from_excel,
    DEFAULT_EXCEL_PATH,
)

DATA_DIR = Path(__file__).parent.parent / "data"
BOM_DIR = DATA_DIR / "saved_boms"
BOM_DIR.mkdir(parents=True, exist_ok=True)

# ============================================================
# Factory Data (from DATA sheet)
# ============================================================
FACTORY_DATA = {
    "KOTOP":            {"workers": 27, "bep": 152, "eff": 0.55,  "epm": 0.11, "region": "VN"},
    "PT.KLATEN":        {"workers": 41, "bep": 140, "eff": 0.512, "epm": 0.10, "region": "IND"},
    "PT.KLATEN Over 15":{"workers": 41, "bep": 140, "eff": 0.385, "epm": 0.10, "region": "IND"},
    "MD":               {"workers": 33, "bep": 230, "eff": 0.70,  "epm": 0.11, "region": "VN"},
}

def calc_cm(smv: float, factory: str) -> float:
    """CM = BEP / (60/SMV * Workers * Efficiency)"""
    if smv <= 0:
        return 0.0
    f = FACTORY_DATA.get(factory, {})
    bep = f.get("bep", 0)
    workers = f.get("workers", 1)
    eff = f.get("eff", 1)
    return round(bep / (60 / smv * workers * eff), 4)


# ============================================================
# Page Config
# ============================================================
st.markdown("""
<style>
    .block-container {
        max-width: 100% !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        padding-top: 2.5rem !important;
    }
    /* Hide sidebar completely */
    section[data-testid="stSidebar"] { display: none !important; }
    button[data-testid="stSidebarCollapsedControl"] { display: none !important; }
    /* Fix left+mid columns, make right col scroll independently */
    [data-testid="stColumns"]:has([data-testid="stRadio"]) {
        height: calc(100vh - 10rem) !important;
        overflow: hidden !important;
    }
    [data-testid="stColumns"]:has([data-testid="stRadio"]) > [data-testid="stColumn"]:nth-child(1),
    [data-testid="stColumns"]:has([data-testid="stRadio"]) > [data-testid="stColumn"]:nth-child(2) {
        height: 100% !important;
        overflow-y: auto !important;
    }
    [data-testid="stColumns"]:has([data-testid="stRadio"]) > [data-testid="stColumn"]:nth-child(3) {
        height: 100% !important;
        overflow-y: auto !important;
    }
    /* Darker table borders */
    [data-testid="stDataFrame"] table,
    [data-testid="stDataEditor"] table { border-collapse: collapse !important; }
    [data-testid="stDataFrame"] th,
    [data-testid="stDataFrame"] td,
    [data-testid="stDataEditor"] th,
    [data-testid="stDataEditor"] td {
        border: 1.5px solid #888 !important;
    }
    [data-testid="stDataFrame"] th,
    [data-testid="stDataEditor"] th {
        border-bottom: 2.5px solid #444 !important;
        background-color: #f0f2f6 !important;
    }
</style>
""", unsafe_allow_html=True)

# Top navigation bar
nav1, nav2, nav3, nav4, nav_spacer = st.columns([1,1,1,1,6])
nav1.page_link("app.py", label="Home")
nav2.page_link("pages/1_PPK_Cost.py", label="**PPK Cost**")
nav3.page_link("pages/2_Color_Cost.py", label="Color Cost")
nav4.page_link("pages/3_Price_Management.py", label="Prices")

# ============================================================
# BOM Loading
# ============================================================
def _load_saved_bom(name):
    path = BOM_DIR / name / "parsed.json"
    if path.exists():
        with open(path, "r", encoding="utf-8") as f:
            return {int(k): v for k, v in json.load(f).items()}
    return None

def _list_saved_boms():
    if not BOM_DIR.exists(): return []
    return sorted([d.name for d in BOM_DIR.iterdir() if d.is_dir() and (d/"parsed.json").exists()], reverse=True)

def _save_bom(name, file_bytes, styles_data):
    dest = BOM_DIR / name
    dest.mkdir(parents=True, exist_ok=True)
    with open(dest / "bom.xlsx", "wb") as f: f.write(file_bytes)
    with open(dest / "parsed.json", "w", encoding="utf-8") as f:
        json.dump(styles_data, f, ensure_ascii=False, indent=2, default=str)


# Top bar: BOM + Excel path (not inside expander to keep state)
saved_boms = _list_saved_boms()
default_mode = 1 if saved_boms else 0  # default to "Load saved" if available
bom_opts = ["Upload new", "Load saved"] if saved_boms else ["Upload new"]

b1, b2, b3 = st.columns([1, 2, 3])
with b1:
    bom_mode = st.radio("BOM", bom_opts, index=default_mode, horizontal=True, key="bom_mode", label_visibility="collapsed")
with b3:
    excel_save_path = st.text_input("Save Path", value=str(DEFAULT_EXCEL_PATH), key="excel_path", label_visibility="collapsed",
                                     placeholder="Costing DB Excel Path")

styles = None
if bom_mode == "Upload new":
    bom_file = st.file_uploader("BOM Extract (.xlsx)", type=["xlsx"], key="bom_upload")
    if bom_file:
        @st.cache_data
        def cached_parse_bom(fb):
            return parse_bom(BytesIO(fb))
        bom_bytes = bom_file.read(); bom_file.seek(0)
        try:
            styles = cached_parse_bom(bom_bytes)
            sc1, sc2 = st.columns([2, 1])
            with sc1:
                bom_name = st.text_input("Save BOM as", value=f"BOM_{datetime.now().strftime('%Y%m%d')}", key="bom_name")
            with sc2:
                if st.button("Save BOM"): _save_bom(bom_name, bom_bytes, styles); st.rerun()
        except Exception as e:
            st.error(f"Error: {e}"); st.stop()
else:
    with b2:
        sel = st.selectbox("BOM", saved_boms, key="sel_bom", label_visibility="collapsed")
    if sel: styles = _load_saved_bom(sel)

if styles is None:
    st.info("Load a BOM to begin."); st.stop()

st.success(f"{len(styles)} styles loaded")

# Load price lists
fabric_prices = load_fabric_prices()
trim_prices = load_trim_prices()

# Dropdown options
def _uq(df, cols, add_empty=True):
    if df.empty: return [""]
    for c in cols:
        if c in df.columns:
            v = sorted(set(str(x).strip() for x in df[c].dropna().unique() if str(x).strip()))
            return ([""] + v) if add_empty else v
    return [""]

fb_item_opts = _uq(fabric_prices, ["ITEM"])
fb_mill_opts = _uq(fabric_prices, ["MILL"])
fb_co_opts = _uq(fabric_prices, ["C/O"])
shrink_opts = ["", "2x1", "3x3", "5x5", "6x6", "7x7", "8x8", "10x10"]
tr_item_opts = _uq(trim_prices, ["ITEM"])
tr_detail_opts = _uq(trim_prices, ["ACC ITEM"])
tr_supplier_opts = _uq(trim_prices, ["Supplier"])
tr_co_opts = _uq(trim_prices, ["C/O"])

# ============================================================
# Cost preview helpers
# ============================================================
def _fb_preview(df):
    rows = []
    for _, r in df.iterrows():
        yy=float(r.get("YY",0)or 0); loss=float(r.get("Loss%",0)or 0)
        su=float(r.get("Size Up%",0)or 0); bp=float(r.get("U/P Buyer",0)or 0)
        np_=float(r.get("U/P Net",0)or 0); tf=float(r.get("Tariff%",0)or 0)
        yf = round(yy*(1+loss)*(1+su),3) if yy>0 else 0
        rows.append({"Placement":r.get("Placement",""), "Item":r.get("Item",""),
            "YY Final":yf, "U/P Buyer":bp, "Cost Buyer":round(yf*bp,3),
            "U/P Net":np_, "Cost Net":round(yf*np_,3), "Tariff%":tf,
            "After Tariff":round(yf*bp*(1-tf),3)})
    return pd.DataFrame(rows)

def _tr_preview(df):
    rows = []
    for _, r in df.iterrows():
        yy=float(r.get("YY",0)or 0); loss=float(r.get("Loss%",0)or 0)
        bp=float(r.get("U/P Buyer",0)or 0); np_=float(r.get("U/P Net",0)or 0)
        tf=float(r.get("Tariff%",0)or 0); rm=float(r.get("RM Air",0)or 0)
        yf = round(yy*(1+loss),3)
        bc=round(yf*bp,3); nc=round(yf*np_,3); tr=round(yf*rm,3)
        rows.append({"Placement":r.get("Placement",""), "Item":r.get("Item",""),
            "YY Final":yf, "U/P Buyer":bp, "Cost Buyer":bc,
            "U/P Net":np_, "Cost Net":nc, "Transit":tr,
            "Total Buyer":round(bc+tr,3), "Total Net":round(nc+tr,3)})
    return pd.DataFrame(rows)

# ============================================================
# MAIN LAYOUT:  Style List | Factory/CM | Detail
# ============================================================
if "style_edits" not in st.session_state:
    st.session_state.style_edits = {}

style_list = sorted(styles.keys())

left_col, mid_col, right_col = st.columns([0.6, 1, 5])

# --- LEFT: Style Selector (narrow + scrollable) ---
with left_col:
    st.markdown("**Styles**")
    # Scrollable container for style list
    style_container = st.container(height=500)
    with style_container:
        selected_style = st.radio(
            "s", style_list, format_func=str,
            key="style_sel", label_visibility="collapsed",
        )

# --- MID: Factory + RM + Size Upcharge ---
with mid_col:
    st.markdown("**Factory**")
    cm_fty = st.selectbox("CM/FTY", list(FACTORY_DATA.keys()), key="cm_fty", label_visibility="collapsed")
    fdata = FACTORY_DATA[cm_fty]
    st.caption(f"Workers **{fdata['workers']}** | BEP **{fdata['bep']}** | Eff **{fdata['eff']}** | Region **{fdata['region']}**")

    st.markdown("**RM Shipmode**")
    rm_shipmode = st.selectbox("RM", ["RM AIR", "RM BOAT"], key="rm_shipmode", label_visibility="collapsed")

    st.markdown("**Size Upcharge**")
    size_upcharge_pct = st.selectbox(
        "SU", [0.1269, 0.0828], key="size_up", label_visibility="collapsed",
        format_func=lambda x: f"{x*100:.2f}% {'(Top)' if x==0.1269 else '(Bottom)'}",
    )

    # Quick CM preview for selected style
    sn = selected_style
    if sn in st.session_state.style_edits:
        _smv = st.session_state.style_edits[sn].get("smv", 0)
        if _smv > 0:
            _cm = calc_cm(_smv, cm_fty)
            st.metric("CM", f"${_cm:.4f}", help=f"SMV {_smv:.2f}")

# --- RIGHT: Selected Style Detail ---
with right_col:
    sn = selected_style
    style_data = styles[sn]

    if sn not in st.session_state.style_edits:
        st.session_state.style_edits[sn] = {
            "smv": 0.0, "wash_price": 0.0, "packing": 0.12,
            "comment": "", "cv_num": "",
        }
    edit = st.session_state.style_edits[sn]

    # --- Version History ---
    saved_versions = list_style_versions_from_excel(sn, excel_save_path)

    # Header row: Style + Version control
    hdr1, hdr2, hdr3 = st.columns([2, 2, 2])
    with hdr1:
        st.subheader(f"Style: {sn}")
    with hdr2:
        save_version = st.text_input("Version", value="1st proto", key=f"ver_{sn}", label_visibility="collapsed",
                                      placeholder="e.g. 1st proto")
    with hdr3:
        if st.button("Save to Excel", key=f"savebtn_{sn}", type="primary", use_container_width=True):
            fb_df = edit.get("fabric_df")
            tr_df = edit.get("trim_df")
            fb_list = fb_df.to_dict("records") if fb_df is not None else []
            tr_list = tr_df.to_dict("records") if tr_df is not None else []
            settings = {"smv": edit.get("smv",0), "wash_price": edit.get("wash_price",0),
                        "packing": edit.get("packing",0.12), "cm_fty": cm_fty,
                        "epm": fdata["epm"], "size_upcharge": size_upcharge_pct,
                        "cv_num": edit.get("cv_num","")}
            save_style_to_excel(sn, save_version, edit.get("comment",""), settings, fb_list, tr_list, excel_save_path)
            save_style(sn, save_version, edit.get("comment",""), settings, fb_list, tr_list)
            st.success(f"Saved!"); st.rerun()

    # Version history
    if saved_versions:
        with st.expander(f"Version History ({len(saved_versions)})", expanded=False):
            for vi, v in enumerate(saved_versions):
                c1,c2,c3,c4 = st.columns([2,3,2,1])
                c1.markdown(f"**{v['version']}**"); c2.caption(v.get("comment","")[:50])
                c3.caption(str(v.get("saved_at","")));
                if c4.button("Load", key=f"ld_{sn}_{vi}"):
                    loaded = load_style_from_excel(sn, v["version"], excel_save_path)
                    if loaded:
                        s = loaded.get("settings",{})
                        edit["smv"]=float(s.get("smv",0)or 0)
                        edit["wash_price"]=float(s.get("wash_price",0)or 0)
                        edit["packing"]=float(s.get("packing",0.12)or 0.12)
                        edit["cv_num"]=s.get("cv_num","")
                        edit["comment"]=loaded.get("comment","")
                        edit["_loaded_fabrics"]=loaded.get("fabrics",[])
                        edit["_loaded_trims"]=loaded.get("trims",[])
                        st.rerun()
            # Copy
            cp1,cp2,cp3 = st.columns([2,2,1])
            copy_from = cp1.selectbox("Copy from", [v["version"] for v in saved_versions], key=f"cpf_{sn}")
            copy_to = cp2.text_input("New name", key=f"cpt_{sn}")
            if cp3.button("Copy", key=f"cpb_{sn}") and copy_to:
                src = load_style_from_excel(sn, copy_from, excel_save_path)
                if src:
                    save_style_to_excel(sn, copy_to, f"Copied from '{copy_from}'. "+src.get("comment",""),
                                        src["settings"], src["fabrics"], src["trims"], excel_save_path)
                    st.rerun()

    # CV# + Comment
    cv_col, cmt_col = st.columns([1, 3])
    with cv_col:
        edit["cv_num"] = st.text_input("CV#", value=edit.get("cv_num",""), key=f"cv_{sn}",
                                        placeholder="e.g. CV-1548900")
    with cmt_col:
        edit["comment"] = st.text_area("Comment", value=edit.get("comment",""), height=50,
                                        placeholder="e.g. 1st proto based on actual sample...", key=f"cmt_{sn}")

    # Settings row: SMV + CM auto-calc + Wash + Packing
    s1,s2,s3,s4 = st.columns(4)
    with s1:
        edit["smv"] = st.number_input("SMV", value=edit["smv"], step=0.01, format="%.2f", key=f"smv_{sn}")
    with s2:
        cm_val = calc_cm(edit["smv"], cm_fty)
        st.metric("CM (auto)", f"${cm_val:.4f}")
        st.caption(f"= {fdata['bep']} / (60/{edit['smv']:.2f} x {fdata['workers']} x {fdata['eff']})" if edit["smv"]>0 else "Enter SMV")
    with s3:
        edit["wash_price"] = st.number_input("Wash Price", value=edit["wash_price"], step=0.01, format="%.3f", key=f"wash_{sn}")
    with s4:
        edit["packing"] = st.number_input("Packing", value=edit["packing"], step=0.01, format="%.3f", key=f"pack_{sn}")

    # ========== FABRIC ==========
    st.markdown("#### Fabric")
    fabric_data = []

    # Priority: 1) version load, 2) previous edit, 3) fresh from BOM
    _loaded_fbs = edit.pop("_loaded_fabrics", None)
    _prev_fb = edit.get("fabric_df")
    if _loaded_fbs:
        for fb in _loaded_fbs:
            fabric_data.append({
                "Placement": fb.get("Placement", fb.get("placement","")),
                "Item": fb.get("Item", fb.get("item","")),
                "Mill": fb.get("Mill", fb.get("mill","")),
                "C/O": fb.get("C/O", fb.get("c_o","")),
                "Shrinkage": fb.get("Shrinkage", fb.get("shrinkage","7x7")),
                "YY": float(fb.get("YY", fb.get("yy",0)) or 0),
                "Loss%": float(fb.get("Loss%", fb.get("loss",0.03)) or 0),
                "Size Up%": float(fb.get("Size Up%", fb.get("size_upcharge", size_upcharge_pct)) or 0),
                "U/P Buyer": float(fb.get("U/P Buyer", fb.get("buyer_price",0)) or 0),
                "U/P Net": float(fb.get("U/P Net", fb.get("net_price",0)) or 0),
                "Tariff%": float(fb.get("Tariff%", fb.get("tariff",0)) or 0),
            })
    elif _prev_fb is not None and not _prev_fb.empty:
        fabric_data = _prev_fb.to_dict("records")
    else:
        for fb in style_data["fabrics"]:
            pi = match_fabric_price(fb["item"], fb["detail"], "", fabric_prices)
            fabric_data.append({
                "Placement": fb["placement"], "Item": fb["item"],
                "Mill": pi["mill"], "C/O": pi["c_o"],
                "Shrinkage": fb.get("shrinkage","7x7"),
                "YY": fb["yy"], "Loss%": fb["loss"], "Size Up%": size_upcharge_pct,
                "U/P Buyer": pi["buyer_price"], "U/P Net": pi["net_price"], "Tariff%": pi["tariff"],
            })

    if fabric_data:
        fabric_df = pd.DataFrame(fabric_data) if not isinstance(fabric_data, pd.DataFrame) else fabric_data
        if isinstance(fabric_data, list):
            fabric_df = pd.DataFrame(fabric_data)
        edited_fabric = st.data_editor(
            fabric_df, key=f"fb_{sn}", use_container_width=True, num_rows="dynamic",
            column_config={
                "Placement": st.column_config.TextColumn("Placement", width="large"),
                "Item": st.column_config.SelectboxColumn("Item", options=fb_item_opts, width="medium"),
                "Mill": st.column_config.SelectboxColumn("Mill", options=fb_mill_opts, width="small"),
                "C/O": st.column_config.SelectboxColumn("C/O", options=fb_co_opts, width="small"),
                "Shrinkage": st.column_config.SelectboxColumn("Shrinkage", options=shrink_opts, width="small"),
                "YY": st.column_config.NumberColumn("YY", format="%.3f"),
                "Loss%": st.column_config.NumberColumn("Loss%", format="%.3f"),
                "Size Up%": st.column_config.NumberColumn("Size Up%", format="%.4f"),
                "U/P Buyer": st.column_config.NumberColumn("U/P Buyer", format="$%.4f"),
                "U/P Net": st.column_config.NumberColumn("U/P Net", format="$%.4f"),
                "Tariff%": st.column_config.NumberColumn("Tariff%", format="%.3f"),
            },
        )
        edit["fabric_df"] = edited_fabric

        # Cost preview
        fbp = _fb_preview(edited_fabric)
        if not fbp.empty and fbp["YY Final"].sum() > 0:
            st.dataframe(
                fbp.style.format({"YY Final":"{:.3f}","U/P Buyer":"${:.4f}","Cost Buyer":"${:.3f}",
                    "U/P Net":"${:.4f}","Cost Net":"${:.3f}","Tariff%":"{:.3f}","After Tariff":"${:.3f}",
                }).map(lambda _:"background-color:#e8f5e9", subset=["Cost Buyer","Cost Net"]),
                use_container_width=True, hide_index=True,
            )
            m1,m2,m3 = st.columns(3)
            m1.metric("Fabric (Buyer)", f"${fbp['Cost Buyer'].sum():.3f}")
            m2.metric("Fabric (Net)", f"${fbp['Cost Net'].sum():.3f}")
            m3.metric("After Tariff", f"${fbp['After Tariff'].sum():.3f}")
    else:
        fbp = pd.DataFrame()

    # ========== TRIM ==========
    st.markdown("#### Trim & Label")
    trim_data = []
    _loaded_trs = edit.pop("_loaded_trims", None)
    _prev_tr = edit.get("trim_df")
    if _loaded_trs:
        for tr in _loaded_trs:
            trim_data.append({
                "Placement": tr.get("Placement",tr.get("placement","")),
                "Item": tr.get("Item",tr.get("item","")),
                "Detail": str(tr.get("Detail",tr.get("detail","")))[:80],
                "Supplier": tr.get("Supplier",tr.get("supplier","")),
                "C/O": tr.get("C/O",tr.get("c_o","")),
                "YY": float(tr.get("YY",tr.get("yy",1.0)) or 0),
                "Loss%": float(tr.get("Loss%",tr.get("loss",0.03)) or 0),
                "U/P Buyer": float(tr.get("U/P Buyer",tr.get("buyer_price",0)) or 0),
                "U/P Net": float(tr.get("U/P Net",tr.get("net_price",0)) or 0),
                "Tariff%": float(tr.get("Tariff%",tr.get("tariff",0)) or 0),
                "RM Air": float(tr.get("RM Air",tr.get("rm_air",0)) or 0),
            })
    elif _prev_tr is not None and not _prev_tr.empty:
        trim_data = _prev_tr.to_dict("records")
    else:
        all_trims = style_data["trims"] + style_data["labels"]
        for tr in all_trims:
            pi = match_trim_price(tr["detail"], trim_prices)
            trim_data.append({
                "Placement": tr["placement"], "Item": tr["item"],
                "Detail": tr["detail"][:80] if tr["detail"] else "",
                "Supplier": pi["supplier"], "C/O": pi["c_o"],
                "YY": tr["yy"], "Loss%": tr["loss"],
                "U/P Buyer": pi["buyer_price"], "U/P Net": pi["net_price"],
                "Tariff%": pi["tariff"], "RM Air": pi["rm_air"],
            })

    if trim_data:
        trim_df = pd.DataFrame(trim_data)
        edited_trim = st.data_editor(
            trim_df, key=f"tr_{sn}", use_container_width=True, num_rows="dynamic",
            column_config={
                "Placement": st.column_config.TextColumn("Placement", width="medium"),
                "Item": st.column_config.SelectboxColumn("Item", options=tr_item_opts, width="small"),
                "Detail": st.column_config.SelectboxColumn("Detail", options=tr_detail_opts, width="large"),
                "Supplier": st.column_config.SelectboxColumn("Supplier", options=tr_supplier_opts, width="small"),
                "C/O": st.column_config.SelectboxColumn("C/O", options=tr_co_opts, width="small"),
                "YY": st.column_config.NumberColumn("YY", format="%.3f"),
                "Loss%": st.column_config.NumberColumn("Loss%", format="%.3f"),
                "U/P Buyer": st.column_config.NumberColumn("U/P Buyer", format="$%.4f"),
                "U/P Net": st.column_config.NumberColumn("U/P Net", format="$%.4f"),
                "Tariff%": st.column_config.NumberColumn("Tariff%", format="%.3f"),
                "RM Air": st.column_config.NumberColumn("RM Air", format="$%.4f"),
            },
        )
        edit["trim_df"] = edited_trim

        trp = _tr_preview(edited_trim)
        if not trp.empty and trp["YY Final"].sum() > 0:
            st.dataframe(
                trp.style.format({"YY Final":"{:.3f}","U/P Buyer":"${:.4f}","Cost Buyer":"${:.3f}",
                    "U/P Net":"${:.4f}","Cost Net":"${:.3f}","Transit":"${:.3f}",
                    "Total Buyer":"${:.3f}","Total Net":"${:.3f}",
                }).map(lambda _:"background-color:#fff3e0", subset=["Total Buyer","Total Net"]),
                use_container_width=True, hide_index=True,
            )
            m1,m2,m3 = st.columns(3)
            m1.metric("Trim (Buyer)", f"${trp['Total Buyer'].sum():.3f}")
            m2.metric("Trim (Net)", f"${trp['Total Net'].sum():.3f}")
    else:
        trp = pd.DataFrame()

    # ========== TOTAL ==========
    st.markdown("---")
    fb_b = fbp["Cost Buyer"].sum() if not fbp.empty else 0
    fb_n = fbp["Cost Net"].sum() if not fbp.empty else 0
    tr_b = trp["Total Buyer"].sum() if not trp.empty else 0
    tr_n = trp["Total Net"].sum() if not trp.empty else 0
    wash = edit.get("wash_price", 0)
    pack = edit.get("packing", 0.12)
    fob_b = round(cm_val + fb_b + tr_b + wash + pack, 2)
    fob_n = round(cm_val + fb_n + tr_n + wash + pack, 2)

    t1,t2,t3,t4,t5,t6 = st.columns(6)
    t1.metric("CM", f"${cm_val:.3f}")
    t2.metric("Fabric (B)", f"${fb_b:.3f}")
    t3.metric("Trim (B)", f"${tr_b:.3f}")
    t4.metric("Wash+Pack", f"${wash+pack:.3f}")
    t5.metric("FOB Buyer", f"${fob_b:.2f}")
    t6.metric("FOB Net", f"${fob_n:.2f}")
