import streamlit as st
import pandas as pd
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from modules.price_matcher import save_price_file, load_fabric_prices, load_trim_prices, load_graphic_prices

DATA_DIR = Path(__file__).parent.parent / "data"

st.markdown("""<style>
section[data-testid="stSidebar"] { display: none !important; }
button[data-testid="stSidebarCollapsedControl"] { display: none !important; }
.block-container { max-width:100%!important; padding-left:1rem!important; padding-right:1rem!important; padding-top:2.5rem!important; }
</style>""", unsafe_allow_html=True)

nav1, nav2, nav3, nav4, _ = st.columns([1,1,1,1,6])
nav1.page_link("app.py", label="Home")
nav2.page_link("pages/1_PPK_Cost.py", label="PPK Cost")
nav3.page_link("pages/2_Color_Cost.py", label="Color Cost")
nav4.page_link("pages/3_Price_Management.py", label="**Prices**")

st.title("Price List Management")
st.markdown("Upload and manage Fabric, Trim, and Graphic price lists. Once uploaded, they are saved internally and used automatically for price matching.")

st.markdown("---")

# --- Fabric Price List ---
st.header("Fabric Price List")
col1, col2 = st.columns([1, 2])

with col1:
    fabric_file = st.file_uploader("Upload Fabric Price List (.xlsx)", type=["xlsx"], key="fabric_upload")
    if fabric_file:
        save_price_file(fabric_file, "fabric")
        st.success("Fabric price list saved!")
        st.rerun()

with col2:
    fabric_df = load_fabric_prices()
    if not fabric_df.empty:
        st.success(f"Loaded: {len(fabric_df)} items")
        st.dataframe(fabric_df, use_container_width=True, height=300)
    else:
        st.info("No fabric price list uploaded yet.")

st.markdown("---")

# --- Trim Price List ---
st.header("Trim Price List")
col1, col2 = st.columns([1, 2])

with col1:
    trim_file = st.file_uploader("Upload Trim Price List (.xlsx)", type=["xlsx"], key="trim_upload")
    if trim_file:
        save_price_file(trim_file, "trim")
        st.success("Trim price list saved!")
        st.rerun()

with col2:
    trim_df = load_trim_prices()
    if not trim_df.empty:
        st.success(f"Loaded: {len(trim_df)} items")
        st.dataframe(trim_df, use_container_width=True, height=300)
    else:
        st.info("No trim price list uploaded yet.")

st.markdown("---")

# --- Graphic Price List ---
st.header("Graphic Price List")
col1, col2 = st.columns([1, 2])

with col1:
    graphic_file = st.file_uploader("Upload Graphic Price List (.xlsx)", type=["xlsx"], key="graphic_upload")
    if graphic_file:
        save_price_file(graphic_file, "graphic")
        st.success("Graphic price list saved!")
        st.rerun()

with col2:
    graphic_df = load_graphic_prices()
    if not graphic_df.empty:
        st.success(f"Loaded: {len(graphic_df)} items")
        st.dataframe(graphic_df, use_container_width=True, height=300)
    else:
        st.info("No graphic price list uploaded yet.")
