import streamlit as st
from pathlib import Path

st.set_page_config(
    page_title="Garment Costing Tool",
    page_icon="",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DATA_DIR = Path(__file__).parent / "data"
DATA_DIR.mkdir(exist_ok=True)
(DATA_DIR / "saved_costs").mkdir(exist_ok=True)

st.markdown("""<style>
section[data-testid="stSidebar"] { display: none !important; }
button[data-testid="stSidebarCollapsedControl"] { display: none !important; }
.block-container { max-width:100%!important; padding-left:1rem!important; padding-right:1rem!important; padding-top:2.5rem!important; }
</style>""", unsafe_allow_html=True)

nav1, nav2, nav3, nav4, _ = st.columns([1,1,1,1,6])
nav1.page_link("app.py", label="**Home**")
nav2.page_link("pages/1_PPK_Cost.py", label="PPK Cost")
nav3.page_link("pages/2_Color_Cost.py", label="Color Cost")
nav4.page_link("pages/3_Price_Management.py", label="Prices")

st.title("Garment Costing Automation")
st.markdown("---")

st.markdown("""
### Welcome!

**BOM Extract** file and upload to automatically generate costing sheets.

#### Pages:
- **PPK Cost** - Stage 1: BOM upload, YY/Loss input, PPK costing
- **Color Cost** - Stage 2: Graphic pricing, final color costing
- **Price Management** - Fabric/Trim/Graphic price list management

Use the sidebar to navigate between pages.
""")

# Show price list status
st.markdown("### Price List Status")
col1, col2, col3 = st.columns(3)

fabric_path = DATA_DIR / "fabric_prices.xlsx"
trim_path = DATA_DIR / "trim_prices.xlsx"
graphic_path = DATA_DIR / "graphic_prices.xlsx"

with col1:
    if fabric_path.exists():
        st.success("Fabric Price List: Loaded")
    else:
        st.warning("Fabric Price List: Not uploaded")

with col2:
    if trim_path.exists():
        st.success("Trim Price List: Loaded")
    else:
        st.warning("Trim Price List: Not uploaded")

with col3:
    if graphic_path.exists():
        st.success("Graphic Price List: Loaded")
    else:
        st.warning("Graphic Price List: Not uploaded")

# Show saved BOMs
st.markdown("### Saved BOM Data")
bom_dir = DATA_DIR / "saved_boms"
if bom_dir.exists():
    saved_boms = [d.name for d in bom_dir.iterdir() if d.is_dir() and (d / "parsed.json").exists()]
    if saved_boms:
        for b in sorted(saved_boms, reverse=True):
            import json
            meta_path = bom_dir / b / "parsed.json"
            with open(meta_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            st.success(f"**{b}** - {len(data)} styles")
    else:
        st.info("No saved BOMs yet. Upload one in PPK Cost page.")
else:
    st.info("No saved BOMs yet. Upload one in PPK Cost page.")

# Show saved costs
saved_dir = DATA_DIR / "saved_costs"
saved_files = list(saved_dir.glob("*.json"))
if saved_files:
    st.markdown(f"### Saved PPK Costs: {len(saved_files)} files")
    for f in sorted(saved_files, reverse=True):
        st.text(f"  {f.stem}")
