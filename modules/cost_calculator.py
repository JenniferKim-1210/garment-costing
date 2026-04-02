import pandas as pd


def calc_fabric_cost(yy: float, loss: float, size_upcharge: float, unit_price: float) -> float:
    """Calculate fabric cost per garment.
    YY_final = YY * (1 + Loss%) * (1 + Size_Upcharge%)
    Cost = YY_final * Unit_Price
    """
    yy_final = round(yy * (1 + loss) * (1 + size_upcharge), 3)
    return round(yy_final * unit_price, 3)


def calc_fabric_yy_final(yy: float, loss: float, size_upcharge: float) -> float:
    return round(yy * (1 + loss) * (1 + size_upcharge), 3)


def calc_trim_cost(yy: float, loss: float, unit_price: float) -> float:
    """Calculate trim cost per garment.
    YY_final = YY * (1 + Loss%)
    Cost = YY_final * Unit_Price
    """
    yy_final = round(yy * (1 + loss), 3)
    return round(yy_final * unit_price, 3)


def calc_trim_yy_final(yy: float, loss: float) -> float:
    return round(yy * (1 + loss), 3)


def calc_after_tariff(cost: float, tariff_pct: float) -> float:
    """Cost after tariff deduction."""
    return round(cost * (1 - tariff_pct), 3)


def calc_transit_cost(yy_final: float, rm_cost_per_yd: float) -> float:
    """RM transit cost = YY_final * RM_cost_per_yard."""
    return round(yy_final * rm_cost_per_yd, 3)


def calc_cm(smv: float, epm: float) -> float:
    """CM = SMV * EPM (earning per minute)."""
    return round(smv * epm, 3)


def calc_cm_from_factory(smv: float, bep: float, workers: int, efficiency: float) -> float:
    """CM = BEP / (60/SMV * Workers * Efficiency)."""
    if smv == 0:
        return 0.0
    pcs_per_hour = 60 / smv
    return round(bep / (pcs_per_hour * workers * efficiency), 3)


def calc_fob(cm: float, fabric_cost: float, trim_cost: float,
             graphic_price: float = 0, wash_price: float = 0,
             packing: float = 0, misc: float = 0) -> float:
    """Calculate total FOB.
    FOB = CM + Fabric Cost + Trim Cost + Graphic + Wash + Packing + Misc
    """
    subtotal = cm + fabric_cost + trim_cost + graphic_price + wash_price + packing
    # Misc is typically calculated as negative percentage of (trim+graphic+wash+packing)
    return round(subtotal + misc, 2)


def calc_misc(trim_cost: float, graphic_price: float, wash_price: float,
              packing: float, misc_pct: float = -0.05) -> float:
    """Misc = -(Trim + Graphic + Wash + Packing) * misc_pct."""
    return round((trim_cost + graphic_price + wash_price + packing) * misc_pct, 3)


def build_style_cost_summary(style_data: dict, fabric_prices_matched: list,
                             trim_prices_matched: list, settings: dict) -> dict:
    """Build a complete cost summary for one style.

    Args:
        style_data: parsed BOM data for this style
        fabric_prices_matched: list of dicts with matched prices + user YY/Loss
        trim_prices_matched: list of dicts with matched prices + user YY/Loss
        settings: dict with cm_fty, smv, packing, wash_price, etc.

    Returns:
        dict with all cost components
    """
    # Fabric costs
    total_fabric_cost = 0.0
    total_fabric_after_tariff = 0.0
    total_fabric_transit = 0.0
    fabric_details = []

    for fb in fabric_prices_matched:
        yy_final = calc_fabric_yy_final(fb["yy"], fb["loss"], fb.get("size_upcharge", 0))
        cost = round(yy_final * fb.get("buyer_price", 0), 3)
        after_tariff = calc_after_tariff(cost, fb.get("tariff", 0))
        transit = calc_transit_cost(yy_final, fb.get("rm_cost", 0))
        total_fabric_cost += cost
        total_fabric_after_tariff += after_tariff
        total_fabric_transit += transit
        fabric_details.append({
            **fb,
            "yy_final": yy_final,
            "cost": cost,
            "after_tariff": after_tariff,
            "transit": transit,
            "total": cost + transit,
        })

    # Trim + Label costs
    total_trim_cost = 0.0
    total_trim_after_tariff = 0.0
    total_trim_transit = 0.0
    trim_details = []

    for tr in trim_prices_matched:
        yy_final = calc_trim_yy_final(tr["yy"], tr["loss"])
        cost = round(yy_final * tr.get("buyer_price", 0), 3)
        after_tariff = calc_after_tariff(cost, tr.get("tariff", 0))
        transit = calc_transit_cost(yy_final, tr.get("rm_cost", 0))
        total_trim_cost += cost
        total_trim_after_tariff += after_tariff
        total_trim_transit += transit
        trim_details.append({
            **tr,
            "yy_final": yy_final,
            "cost": cost,
            "after_tariff": after_tariff,
            "transit": transit,
            "total": cost + transit,
        })

    # CM
    smv = settings.get("smv", 0)
    epm = settings.get("epm", 0.11)
    cm = calc_cm(smv, epm)

    # Other costs
    graphic_price = settings.get("graphic_price", 0)
    wash_price = settings.get("wash_price", 0)
    packing = settings.get("packing", 0)

    misc = calc_misc(total_trim_cost + total_trim_transit,
                     graphic_price, wash_price, packing,
                     settings.get("misc_pct", -0.05))

    fob = calc_fob(
        cm=cm,
        fabric_cost=total_fabric_cost + total_fabric_transit,
        trim_cost=total_trim_cost + total_trim_transit,
        graphic_price=graphic_price,
        wash_price=wash_price,
        packing=packing,
        misc=misc,
    )

    return {
        "cm": cm,
        "smv": smv,
        "fabric_cost": total_fabric_cost,
        "fabric_transit": total_fabric_transit,
        "fabric_total": total_fabric_cost + total_fabric_transit,
        "fabric_after_tariff": total_fabric_after_tariff,
        "trim_cost": total_trim_cost,
        "trim_transit": total_trim_transit,
        "trim_total": total_trim_cost + total_trim_transit,
        "trim_after_tariff": total_trim_after_tariff,
        "graphic_price": graphic_price,
        "wash_price": wash_price,
        "packing": packing,
        "misc": misc,
        "fob": fob,
        "fabric_details": fabric_details,
        "trim_details": trim_details,
    }
