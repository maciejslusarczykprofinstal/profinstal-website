from flask import Flask, render_template_string, request, jsonify

app = Flask(__name__)

EXPERT = {
    "name": "mgr inż. Maciej Ślusarczyk",
    "title": "Ekspert HVAC / instalacje sanitarne",
    "lic": "Uprawnienia budowlane bez ograniczeń, nr XXX/XX/XX",
    "chamber": "Członek Małopolskiej OIIB",
    "contact": "kontakt@profinstal.info | +48 123 456 789",
    "company": "PROF INSTAL",
    "city": "Kraków"
}

CP_KJ_PER_KG_K = 4.19  # kJ/(kg·K)

def price_GJ_brutto(price_net: float, unit: str, vat_percent: float) -> float:
    per_GJ_net = price_net * 1000.0 if unit.upper() == "MJ" else price_net
    return per_GJ_net * (1.0 + vat_percent/100.0)

def Q_GJ_per_m3(dT: float) -> float:
    # Q = m·c·ΔT; m=1000 kg; c=4.19 kJ/kgK; GJ = kJ / 1e6
    return (1000.0 * CP_KJ_PER_KG_K * dT) / 1_000_000.0

def compute(payload: dict) -> dict:
    bill       = float(payload["bill"])             # zł/m³ (rachunek)
    heat_price = float(payload["heat_price"])       # zł/GJ lub zł/MJ (netto)
    unit       = payload.get("unit","GJ").upper()   # "GJ" / "MJ"
    vat        = float(payload.get("vat","23"))
    month_m3   = float(payload["month_m3"])
    units      = int(payload["units"])
    dT         = float(payload["dT"])

    c_gj_brutto = price_GJ_brutto(heat_price, unit, vat)
    q_per_m3    = Q_GJ_per_m3(dT)
    cost_theor  = q_per_m3 * c_gj_brutto
    eta         = max(min(cost_theor / bill, 1.0), 0.0) if bill > 0 else 0.0

    loss_per_m3  = bill - cost_theor
    loss_flat_m  = loss_per_m3 * month_m3
    loss_build_m = loss_flat_m * units
    loss_build_y = loss_build_m * 12.0

    def cost_at_eff(eff): return (q_per_m3 / eff) * c_gj_brutto
    cost70 = cost_at_eff(0.70); cost80 = cost_at_eff(0.80)
    save70_m3 = max(bill - cost70, 0.0); save80_m3 = max(bill - cost80, 0.0)
    save70_flat_m = save70_m3 * month_m3; save80_flat_m = save80_m3 * month_m3
    save70_build_m = save70_flat_m * units; save80_build_m = save80_flat_m * units
    save70_build_y = save70_build_m * 12.0; save80_build_y = save80_build_m * 12.0

    return {
        "bill": bill, "heat_price": heat_price, "unit": unit, "vat": vat,
        "dT": dT, "month_m3": month_m3, "units": units,
        "price_GJ_brutto": c_gj_brutto, "q_per_m3": q_per_m3,
        "cost_theor": cost_theor, "eta": eta,
        "loss_per_m3": loss_per_m3, "loss_flat_m": loss_flat_m,
        "loss_build_m": loss_build_m, "loss_build_y": loss_build_y,
        "cost70": cost70, "cost80": cost80,
        "save70_m3": save70_m3, "save80_m3": save80_m3,
        "save70_build_m": save70_build_m, "save80_build_m": save80_build_m,
        "save70_build_y": save70_build_y, "save80_build_y": save80_build_y
    }

PAGE = """
<!doctype html><html lang="pl"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>PROF INSTAL – Kalkulator CWU (web)</title>
<style>
:root{--bg:#0b1220;--surface:#0f172a;--line:#1e293b;--text:#e2e8f0;--muted:#94a3b8;--brand:#10b981}
*{box-sizing:border-box} body{margin:0;font-family:system-ui,Segoe UI,Roboto,Inter,Arial;background:var(--bg);color:var(--text)}
.wrap{max-width:1024px;margin:auto;padding:20px}
.card{background:linear-gradient(180deg,rgba(255,255,255,.04),rgba(255,255,255,.02));border:1px solid var(--line);border-radius:18px;padding:16px}
h1{margin:8px 0 12px 0} .muted{color:var(--muted)}
.grid{display:grid;gap:12px} .g2{grid-template-columns:1fr 1fr} .g4{grid-template-columns:repeat(4,1fr)}
@media (max-width:900px){.g2,.g4{grid-template-columns:1fr}}
.input, select{width:100%;padding:10px;border-radius:10px;border:1px solid var(--line);background:#0b1730;color:var(--text)}
.btn{border:0;border-radius:12px;padding:12px 16px;font-weight:700;background:var(--brand);color:#01261f;cursor:pointer}
.kpi{border:1px solid var(--line);border-radius:14px;padding:14px} .kpi h3{margin:0 0 6px 0}
.code{font-family:ui-monospace,Consolas,monospace;background:#0a1428;padding:8px;border-radius:8px;border:1px solid var(--line)}
</style></head><body>
<div class="wrap">
  <div class="card">
    <h1>PROF INSTAL — Kalkulator CWU (web)</h1>
    <p class="muted">{{expert.company}} • {{expert.name}} — {{expert.title}} • {{expert.city}}</p>
    <form method="post" class="grid g4" action="/">
      <div><label>Stawka z rachunku [zł/m³]</label><input class="input" name="bill" value="{{f.bill}}" required></div>
      <div><label>Cena ciepła (netto)</label><input class="input" name="heat_price" value="{{f.heat_price}}" required></div>
      <div><label>Jednostka ceny</label>
        <select class="input" name="unit">
          <option value="GJ" {% if f.unit=='GJ' %}selected{% endif %}>zł/GJ</option>
          <option value="MJ" {% if f.unit=='MJ' %}selected{% endif %}>zł/MJ</option>
        </select>
      </div>
      <div><label>VAT [%]</label><input class="input" name="vat" value="{{f.vat}}" required></div>

      <div><label>ΔT podgrzewu [°C]</label><input class="input" name="dT" value="{{f.dT}}" required></div>
      <div><label>Zużycie / mies. [m³]</label><input class="input" name="month_m3" value="{{f.month_m3}}" required></div>
      <div><label>Liczba mieszkań w budynku</label><input class="input" name="units" value="{{f.units}}" required></div>
      <div style="display:flex;align-items:flex-end"><button class="btn" type="submit">Oblicz</button></div>
    </form>
  </div>

  {% if r %}
  <div class="grid g2" style="margin-top:12px">
    <div class="card">
      <h3>Wyniki — fizyka i koszty</h3>
      <div class="grid g2">
        <div class="kpi"><h3>{{ "%.5f"|format(r.q_per_m3) }} GJ/m³</h3><div class="muted">Q_teor na 1 m³</div></div>
        <div class="kpi"><h3>{{ "%.2f"|format(r.price_GJ_brutto) }} zł/GJ</h3><div class="muted">Cena ciepła brutto</div></div>
        <div class="kpi"><h3>{{ "%.2f"|format(r.cost_theor) }} zł/m³</h3><div class="muted">Koszt teoretyczny</div></div>
        <div class="kpi"><h3>{{ "%.1f"|format(r.eta*100) }} %</h3><div class="muted">Sprawność η</div></div>
        <div class="kpi"><h3>{{ "%.2f"|format(r.loss_per_m3) }} zł/m³</h3><div class="muted">Strata na 1 m³</div></div>
        <div class="kpi"><h3>{{ "%.2f"|format(r.loss_flat_m) }} zł/mies</h3><div class="muted">Strata lokalu</div></div>
        <div class="kpi"><h3>{{ "{:,.2f}".format(r.loss_build_m) }} zł/mies</h3><div class="muted">Strata budynku</div></div>
        <div class="kpi"><h3>{{ "{:,.2f}".format(r.loss_build_y) }} zł/rok</h3><div class="muted">Strata budynku rocznie</div></div>
      </div>
    </div>
    <div class="card">
      <h3>Scenariusze modernizacji</h3>
      <div class="grid g2">
        <div class="kpi"><h3>70% → {{ "%.2f"|format(r.cost70) }} zł/m³</h3><div class="muted">koszt / oszcz. {{ "%.2f"|format(r.save70_m3) }} zł/m³</div></div>
        <div class="kpi"><h3>80% → {{ "%.2f"|format(r.cost80) }} zł/m³</h3><div class="muted">koszt / oszcz. {{ "%.2f"|format(r.save80_m3) }} zł/m³</div></div>
        <div class="kpi"><h3>{{ "{:,.2f}".format(r.save70_build_m) }} zł/mies</h3><div class="muted">oszcz. budynku (70%)</div></div>
        <div class="kpi"><h3>{{ "{:,.2f}".format(r.save80_build_m) }} zł/mies</h3><div class="muted">oszcz. budynku (80%)</div></div>
        <div class="kpi"><h3>{{ "{:,.0f}".format(r.save70_build_y) }} zł/rok</h3><div class="muted">oszcz. budynku (70%)</div></div>
        <div class="kpi"><h3>{{ "{:,.0f}".format(r.save80_build_y) }} zł/rok</h3><div class="muted">oszcz. budynku (80%)</div></div>
      </div>
      <p class="muted" style="margin-top:8px">η = koszt_teor / stawka_rachunkowa (im bliżej 100%, tym lepiej).</p>
    </div>
  </div>
  {% endif %}

  <div class="card" style="margin-top:12px">
    <h3>Wzory</h3>
    <pre class="code">Q = m·c·ΔT;   m=1000 kg,  c=4.19 kJ/(kg·K)
Q[GJ/m³] = (1000·4.19·ΔT) / 1e6
koszt_teor [zł/m³] = Q · cena_ciepła_brutto [zł/GJ]
η = koszt_teor / stawka_rachunkowa</pre>
  </div>
</div>
</body></html>
"""

@app.route("/", methods=["GET","POST"])
def index():
    defaults = dict(bill="49.00", heat_price="73.69", unit="GJ", vat="23", dT="45", month_m3="7.42", units="65")
    result = None
    form_vals = defaults.copy()
    if request.method == "POST":
        form_vals = {k: request.form.get(k, defaults[k]) for k in defaults.keys()}
        try:
            result = compute(form_vals)
        except Exception as e:
            result = None
    return render_template_string(PAGE, expert=EXPERT, r=result, f=form_vals)

@app.post("/api/calc")
def api_calc():
    data = request.get_json(force=True)
    out = compute(data)
    return jsonify(out)

if __name__ == "__main__":
    app.run(debug=True, host="127.0.0.1", port=5000)
