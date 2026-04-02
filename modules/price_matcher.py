import pandas as pd
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"


def load_fabric_prices(file_path=None) -> pd.DataFrame:
    """Load fabric price list. Uses internal DB if no file provided."""
    if file_path is None:
        file_path = DATA_DIR / "fabric_prices.xlsx"
    if not Path(file_path).exists():
        return pd.DataFrame()
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_trim_prices(file_path=None) -> pd.DataFrame:
    """Load trim price list."""
    if file_path is None:
        file_path = DATA_DIR / "trim_prices.xlsx"
    if not Path(file_path).exists():
        return pd.DataFrame()
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_graphic_prices(file_path=None) -> pd.DataFrame:
    """Load graphic price list."""
    if file_path is None:
        file_path = DATA_DIR / "graphic_prices.xlsx"
    if not Path(file_path).exists():
        return pd.DataFrame()
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def save_price_file(uploaded_file, price_type: str):
    """Save uploaded price file to internal data directory."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    filename_map = {
        "fabric": "fabric_prices.xlsx",
        "trim": "trim_prices.xlsx",
        "graphic": "graphic_prices.xlsx",
    }
    dest = DATA_DIR / filename_map[price_type]
    with open(dest, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return dest


def match_fabric_price(fabric_item: str, fabric_detail: str, color: str,
                       fabric_prices: pd.DataFrame) -> dict:
    """Match a fabric item to its price from the price list.

    Returns dict with keys: buyer_price, net_price, mill, c_o, tariff, color_type, matched
    """
    result = {
        "buyer_price": 0.0, "net_price": 0.0, "mill": "", "c_o": "",
        "tariff": 0.0, "color_type": "", "matched": False, "cw": 0, "weight": 0,
    }
    if fabric_prices.empty:
        return result

    # Column name mapping (handle variations)
    item_col = _find_col(fabric_prices, ["ITEM", "FABRIC ITEM", "Item"])
    detail_col = _find_col(fabric_prices, ["FABRIC ITEM", "Detail"])
    buyer_col = _find_col(fabric_prices, ["BUYER", "Buyer Cost", "H"])
    net_col = _find_col(fabric_prices, ["NET", "Net Cost", "I"])
    mill_col = _find_col(fabric_prices, ["MILL", "E"])
    co_col = _find_col(fabric_prices, ["C/O", "F"])
    tariff_col = _find_col(fabric_prices, ["TARIFF %", "TARIFF", "G"])
    color_col = _find_col(fabric_prices, ["COLOR", "D"])
    cw_col = _find_col(fabric_prices, ["CW", "J"])
    weight_col = _find_col(fabric_prices, ["WEIGHT", "L"])

    if item_col is None:
        return result

    # Try matching by ITEM name first
    for idx, row in fabric_prices.iterrows():
        item_val = str(row.get(item_col, "")).strip()
        if item_val and fabric_item and _fuzzy_match(fabric_item, item_val):
            result["buyer_price"] = _safe_float(row.get(buyer_col, 0))
            result["net_price"] = _safe_float(row.get(net_col, 0))
            result["mill"] = str(row.get(mill_col, "")).strip() if mill_col else ""
            result["c_o"] = str(row.get(co_col, "")).strip() if co_col else ""
            result["tariff"] = _safe_float(row.get(tariff_col, 0))
            result["color_type"] = str(row.get(color_col, "")).strip() if color_col else ""
            result["cw"] = _safe_float(row.get(cw_col, 0))
            result["weight"] = _safe_float(row.get(weight_col, 0))
            result["matched"] = True
            break

    return result


def match_trim_price(trim_detail: str, trim_prices: pd.DataFrame) -> dict:
    """Match a trim item to its price from the price list.

    Returns dict with keys: buyer_price, net_price, supplier, c_o, tariff, matched
    """
    result = {
        "buyer_price": 0.0, "net_price": 0.0, "supplier": "", "c_o": "",
        "tariff": 0.0, "matched": False, "rm_air": 0.0,
    }
    if trim_prices.empty or not trim_detail:
        return result

    acc_col = _find_col(trim_prices, ["ACC ITEM", "B"])
    buyer_col = _find_col(trim_prices, ["Buyer Cost - ACC U/P", "Buyer Cost", "F"])
    net_col = _find_col(trim_prices, ["NET Cost - ACC U/P", "NET Cost", "Net Cost", "G"])
    supplier_col = _find_col(trim_prices, ["Supplier", "C"])
    co_col = _find_col(trim_prices, ["C/O", "D"])
    tariff_col = _find_col(trim_prices, ["TARIFF %", "TARIFF", "H"])
    rm_col = _find_col(trim_prices, ["RM AIR", "J"])

    if acc_col is None:
        return result

    best_score = 0
    for idx, row in trim_prices.iterrows():
        acc_val = str(row.get(acc_col, "")).strip()
        if not acc_val:
            continue
        score = _match_score(trim_detail, acc_val)
        if score > best_score and score > 0.5:
            best_score = score
            result["buyer_price"] = _safe_float(row.get(buyer_col, 0))
            result["net_price"] = _safe_float(row.get(net_col, 0))
            result["supplier"] = str(row.get(supplier_col, "")).strip() if supplier_col else ""
            result["c_o"] = str(row.get(co_col, "")).strip() if co_col else ""
            result["tariff"] = _safe_float(row.get(tariff_col, 0))
            result["rm_air"] = _safe_float(row.get(rm_col, 0))
            result["matched"] = True

    return result


def match_graphic_price(graphic_code: str, colorway_code: str,
                        graphic_prices: pd.DataFrame) -> dict:
    """Match a graphic code to its price."""
    result = {
        "buyer_price": 0.0, "net_price": 0.0, "technique": "",
        "placement_size": "", "matched": False,
    }
    if graphic_prices.empty or not graphic_code:
        return result

    code_col = _find_col(graphic_prices, ["GRAPHIC CODE", "A"])
    cw_col = _find_col(graphic_prices, ["COLORWAY", "B"])
    buyer_col = _find_col(graphic_prices, ["BUYER COST", "F"])
    net_col = _find_col(graphic_prices, ["NET COST", "G"])
    tech_col = _find_col(graphic_prices, ["TECHNIQUE", "D"])
    place_col = _find_col(graphic_prices, ["PLACEMENT & SIZE", "E"])

    if code_col is None:
        return result

    # Extract base graphic code (e.g., "GPH26-078" from "GPH26-078 C002")
    base_code = graphic_code.split()[0] if graphic_code else ""

    for idx, row in graphic_prices.iterrows():
        row_code = str(row.get(code_col, "")).strip()
        if row_code and base_code and row_code == base_code:
            result["buyer_price"] = _safe_float(row.get(buyer_col, 0))
            result["net_price"] = _safe_float(row.get(net_col, 0))
            result["technique"] = str(row.get(tech_col, "")).strip() if tech_col else ""
            result["placement_size"] = str(row.get(place_col, "")).strip() if place_col else ""
            result["matched"] = True
            # If colorway column exists, try more specific match
            if cw_col:
                row_cw = str(row.get(cw_col, "")).strip()
                if colorway_code and colorway_code in row_cw:
                    break  # Exact colorway match, stop
            else:
                break

    return result


def _find_col(df: pd.DataFrame, candidates: list) -> str | None:
    """Find the first matching column name from candidates."""
    for c in candidates:
        if c in df.columns:
            return c
    for c in candidates:
        for col in df.columns:
            if c.lower() in col.lower():
                return col
    return None


def _fuzzy_match(query: str, target: str) -> bool:
    """Simple fuzzy matching for fabric/trim items."""
    q = query.lower().strip()
    t = target.lower().strip()
    return q == t or q in t or t in q


def _match_score(query: str, target: str) -> float:
    """Calculate match score between two strings."""
    q = query.lower().strip()
    t = target.lower().strip()
    if q == t:
        return 1.0
    if q in t or t in q:
        return 0.8

    # Word overlap
    q_words = set(q.split())
    t_words = set(t.split())
    if not q_words or not t_words:
        return 0.0
    overlap = len(q_words & t_words)
    return overlap / max(len(q_words), len(t_words))


def _safe_float(val) -> float:
    try:
        if pd.isna(val):
            return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0
