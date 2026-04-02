import streamlit as st
import pandas as pd
import json
from pathlib import Path
from io import BytesIO
from datetime import datetime
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from modules.price_matcher import load_graphic_prices, match_graphic_price
from modules.cost_calculator import calc_fob, calc_misc
from modules.summary_writer import create_summary_excel

DATA_DIR = Path(__file__).parent.parent / "data"

st.markdown("""<style>
section[data-testid="stSidebar"] { display: none !important; }
button[data-testid="stSidebarCollapsedControl"] { display: none !important; }
.block-container { max-width:100%!important; padding-left:1rem!important; padding-right:1rem!important; padding-top:2.5rem!important; }
</style>""", unsafe_allow_html=True)

nav1, nav2, nav3, nav4, _ = st.columns([1,1,1,1,6])
nav1.page_link("app.py", label="Home")
nav2.page_link("pages/1_PPK_Cost.py", label="PPK Cost")
nav3.page_link("pages/2_Color_Cost.py", label="**Color Cost**")
nav4.page_link("pages/3_Price_Management.py", label="Prices")

st.title("Stage 2: Color Cost")
st.markdown("Load saved PPK Cost -> Add Graphic pricing -> Generate final color costing")

# ============================================================
# Step 1: Load Saved PPK Cost
# ============================================================
st.header("Step 1: Load PPK Cost")

saved_dir = DATA_DIR / "saved_costs"
saved_files = sorted(saved_dir.glob("*.json"), reverse=True) if saved_dir.exists() else []

if not saved_files:
    st.warning("No saved PPK costs found. Please complete Stage 1 (PPK Cost) first.")
    st.stop()

selected_file = st.selectbox(
    "Select saved PPK Cost",
    options=saved_files,
    format_func=lambda x: x.stem,
)

if selected_file:
    with open(selected_file, "r", encoding="utf-8") as f:
        ppk_data = json.load(f)
    st.success(f"Loaded {len(ppk_data)} styles from {selected_file.stem}")

    # Show summary of loaded data
    summary_rows = []
    for sn, data in ppk_data.items():
        summary_rows.append({
            "Style#": sn,
            "CM": data.get("cm", 0),
            "Fabric Total": data.get("fabric_total", 0),
            "Trim Total": data.get("trim_total", 0),
            "PPK FOB (no graphic)": data.get("fob", 0),
        })
    st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

# ============================================================
# Step 2: Graphic Matching
# ============================================================
st.header("Step 2: Graphic Pricing")

graphic_prices = load_graphic_prices()
if graphic_prices.empty:
    st.warning("No Graphic price list loaded. Go to Price Management to upload.")
    graphic_prices = pd.DataFrame()

st.markdown("#### Assign Graphic Code & Price per Style")

if "graphic_edits" not in st.session_state:
    st.session_state.graphic_edits = {}

graphic_rows = []
for sn, data in ppk_data.items():
    existing = st.session_state.graphic_edits.get(sn, {})
    graphic_rows.append({
        "Style#": sn,
        "Graphic Code": existing.get("graphic_code", ""),
        "Technique": existing.get("technique", ""),
        "Buyer Cost": existing.get("buyer_cost", 0.0),
        "Net Cost": existing.get("net_cost", 0.0),
        "PPK FOB": data.get("fob", 0),
    })

graphic_df = pd.DataFrame(graphic_rows)
edited_graphic = st.data_editor(
    graphic_df,
    key="graphic_editor",
    use_container_width=True,
    column_config={
        "Style#": st.column_config.TextColumn("Style#", disabled=True),
        "Graphic Code": st.column_config.TextColumn("Graphic Code"),
        "Technique": st.column_config.TextColumn("Technique"),
        "Buyer Cost": st.column_config.NumberColumn("Buyer Cost", format="%.4f"),
        "Net Cost": st.column_config.NumberColumn("Net Cost", format="%.4f"),
        "PPK FOB": st.column_config.NumberColumn("PPK FOB", format="%.2f", disabled=True),
    },
)

# Auto-match button
if not graphic_prices.empty:
    if st.button("Auto-Match Graphic Prices"):
        for idx, row in edited_graphic.iterrows():
            code = row.get("Graphic Code", "")
            if code:
                match = match_graphic_price(code, "", graphic_prices)
                if match["matched"]:
                    edited_graphic.at[idx, "Buyer Cost"] = match["buyer_price"]
                    edited_graphic.at[idx, "Net Cost"] = match["net_price"]
                    edited_graphic.at[idx, "Technique"] = match["technique"]
        st.rerun()

# ============================================================
# Step 3: Calculate Final Color Cost
# ============================================================
st.header("Step 3: Final Color Cost")

if st.button("Calculate Final FOB with Graphics", type="primary", use_container_width=True):
    final_costs = {}

    for idx, row in edited_graphic.iterrows():
        sn = str(row["Style#"])
        data = ppk_data.get(sn, {})

        graphic_buyer = float(row.get("Buyer Cost", 0) or 0)

        # Recalculate FOB with graphic
        cm = data.get("cm", 0)
        fabric_total = data.get("fabric_total", 0)
        trim_total = data.get("trim_total", 0)
        wash_price = data.get("wash_price", 0)
        packing = data.get("packing", 0)

        misc = calc_misc(trim_total, graphic_buyer, wash_price, packing)
        fob = calc_fob(cm, fabric_total, trim_total, graphic_buyer, wash_price, packing, misc)

        final_costs[sn] = {
            **data,
            "graphic_price": graphic_buyer,
            "graphic_code": row.get("Graphic Code", ""),
            "technique": row.get("Technique", ""),
            "misc": misc,
            "fob": fob,
            "fob_ppk": data.get("fob", 0),
        }

    # Show final summary
    st.markdown("### Final Color Cost Summary")
    final_rows = []
    for sn, data in final_costs.items():
        final_rows.append({
            "Style#": sn,
            "CM": data.get("cm", 0),
            "Fabric": data.get("fabric_total", 0),
            "Trim": data.get("trim_total", 0),
            "Graphic": data.get("graphic_price", 0),
            "Wash": data.get("wash_price", 0),
            "Packing": data.get("packing", 0),
            "Misc": data.get("misc", 0),
            "PPK FOB": data.get("fob_ppk", 0),
            "Final FOB": data.get("fob", 0),
            "Diff": round(data.get("fob", 0) - data.get("fob_ppk", 0), 3),
        })

    final_df = pd.DataFrame(final_rows)
    st.dataframe(final_df, use_container_width=True)

    # Highlight graphic impact
    total_graphic = sum(d.get("graphic_price", 0) for d in final_costs.values())
    st.info(f"Total Graphic Cost Impact: ${total_graphic:.2f}")

    # Download
    summary_bytes = create_summary_excel(final_costs)
    st.download_button(
        "Download Final Color Cost Excel",
        data=summary_bytes,
        file_name=f"Color_Cost_Final_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # Save final
    if st.button("Save Final Color Cost"):
        save_path = DATA_DIR / "saved_costs" / f"color_cost_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(save_path, "w", encoding="utf-8") as f:
            json.dump(final_costs, f, ensure_ascii=False, indent=2, default=str)
        st.success(f"Saved to {save_path.name}")
