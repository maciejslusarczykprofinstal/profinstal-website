
# -*- coding: utf-8 -*-
"""
PROF INSTAL — core calculation logic extracted from the Tkinter app
© 2025 Maciej Ślusarczyk. All rights reserved.
"""
from math import isfinite

CP_KJ_PER_KG_K = 4.19  # kJ/(kg·K)

def price_GJ_brutto(price_net: float, unit: str = "GJ", vat_percent: float = 23.0) -> float:
    """
    Convert heat price to PLN/GJ including VAT.
    unit: "GJ" (zł/GJ) or "MJ" (zł/MJ)
    """
    per_GJ_net = price_net * 1000.0 if unit.upper() == "MJ" else price_net
    return per_GJ_net * (1.0 + float(vat_percent)/100.0)

def Q_GJ_per_m3(dT: float) -> float:
    """
    Q = m*c*ΔT; m=1000 kg, c=4.19 kJ/kgK; convert kJ to GJ (1e6)
    """
    return (1000.0 * CP_KJ_PER_KG_K * dT) / 1_000_000.0

def compute_all(bill: float, heat_price: float, unit: str, vat: float, month_m3: float, units: int, dT: float):
    # sanitize
    if not all(isfinite(x) for x in [bill, heat_price, vat, month_m3, dT]) or units <= 0:
        raise ValueError("Invalid inputs")
    unit = unit.upper()
    price_gj_brutto = price_GJ_brutto(heat_price, unit, vat)
    q_per_m3 = Q_GJ_per_m3(dT)
    cost_theor = q_per_m3 * price_gj_brutto
    eta = max(min(cost_theor / bill if bill != 0 else 0.0, 1.0), 0.0)

    loss_per_m3 = bill - cost_theor
    loss_flat_m = loss_per_m3 * month_m3
    loss_build_m = loss_flat_m * units
    loss_build_y = loss_build_m * 12.0

    def cost_at_eff(eff): 
        return (q_per_m3 / eff) * price_gj_brutto

    cost70 = cost_at_eff(0.70); cost80 = cost_at_eff(0.80)
    save70_m3 = max(bill - cost70, 0.0); save80_m3 = max(bill - cost80, 0.0)
    save70_flat_m = save70_m3 * month_m3; save80_flat_m = save80_m3 * month_m3
    save70_build_m = save70_flat_m * units; save80_build_m = save80_flat_m * units
    save70_build_y = save70_build_m * 12.0; save80_build_y = save80_build_m * 12.0

    return {
        "bill": bill, "heat_price": heat_price, "unit": unit, "vat": vat, "dT": dT, "month_m3": month_m3, "units": units,
        "price_GJ_brutto": price_gj_brutto, "q_per_m3": q_per_m3, "cost_theor": cost_theor, "eta": eta,
        "loss_per_m3": loss_per_m3, "loss_flat_m": loss_flat_m, "loss_build_m": loss_build_m, "loss_build_y": loss_build_y,
        "cost70": cost70, "cost80": cost80, "save70_m3": save70_m3, "save80_m3": save80_m3,
        "save70_flat_m": save70_flat_m, "save80_flat_m": save80_flat_m,
        "save70_build_m": save70_build_m, "save80_build_m": save80_build_m,
        "save70_build_y": save70_build_y, "save80_build_y": save80_build_y
    }


# --- AUDYTORSKIE PORÓWNANIE INSTALACJI ---
def compute_audit(params_old: dict, params_new: dict) -> dict:
    """
    Porównuje straty starej i nowej instalacji na podstawie parametrów i wzorów z norm.
    params_old, params_new: dict z kluczami:
        - 'Q' (moc [kW]), 'L' (długość przewodów [m]), 'd' (średnica [mm]),
        - 'lambda' (wsp. przewodzenia [W/mK]), 't_in' (temp. zasilania [°C]),
        - 't_out' (temp. powrotu [°C]), 't_amb' (temp. otoczenia [°C]),
        - 'ins_thick' (grubość izolacji [mm]), 'czas_pracy' (h/rok)
    Zwraca słownik z porównaniem strat i oszczędności.
    """
    def heat_loss(Q, L, d, lamb, t_in, t_out, t_amb, ins_thick, czas_pracy):
        # Wzór uproszczony wg PN-EN 12831, PN-EN 15316
        # Strata liniowa: q = 2 * pi * lambda * (t_sr - t_amb) / ln((d+2*ins)/d)
        from math import pi, log
        d_m = d / 1000.0
        ins_m = ins_thick / 1000.0
        d_ext = d_m + 2*ins_m
        t_sr = (t_in + t_out) / 2.0
        q = 2 * pi * lamb * (t_sr - t_amb) / log(d_ext/d_m)
        Q_loss = q * L * czas_pracy / 1000.0  # [kWh/rok]
        return Q_loss

    Q_loss_old = heat_loss(
        params_old['Q'], params_old['L'], params_old['d'], params_old['lambda'],
        params_old['t_in'], params_old['t_out'], params_old['t_amb'], params_old['ins_thick'], params_old['czas_pracy']
    )
    Q_loss_new = heat_loss(
        params_new['Q'], params_new['L'], params_new['d'], params_new['lambda'],
        params_new['t_in'], params_new['t_out'], params_new['t_amb'], params_new['ins_thick'], params_new['czas_pracy']
    )
    oszczednosc = Q_loss_old - Q_loss_new
    procent = 100.0 * oszczednosc / Q_loss_old if Q_loss_old else 0.0
    return {
        'Q_loss_old': Q_loss_old,
        'Q_loss_new': Q_loss_new,
        'oszczednosc_kWh': oszczednosc,
        'oszczednosc_proc': procent
    }
