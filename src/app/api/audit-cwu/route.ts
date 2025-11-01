import { NextResponse } from "next/server";

export async function POST(req: Request) {
  try {
    const data = await req.json();
    const errs: string[] = [];

    const need = ["E_DHW_GJ","V_DHW_m3","price_GJ_PLN","T_hot_C","T_cold_C","period"];
    for (const k of need) if (data[k] === undefined) errs.push(`${k} jest wymagane`);
    if (errs.length) return NextResponse.json({ ok:false, errors:errs }, { status:400 });

    const E = Number(data.E_DHW_GJ);
    const V = Number(data.V_DHW_m3);
    const P = Number(data.price_GJ_PLN);
    const Th = Number(data.T_hot_C ?? 55);
    const Tc = Number(data.T_cold_C ?? 8);
    if (!(E>0)) errs.push("E_DHW_GJ > 0");
    if (!(V>0)) errs.push("V_DHW_m3 > 0");
    if (!(P>0)) errs.push("price_GJ_PLN > 0");
    if (!(Th > Tc)) errs.push("T_hot_C > T_cold_C");
    if (errs.length) return NextResponse.json({ ok:false, errors:errs }, { status:400 });

    const deltaT = Th - Tc;
    const E_th = 0.004186 * deltaT;   // GJ/m3
    const E_real = E / V;             // GJ/m3
    const eta = E_th / E_real;
    const loss_pct = (1 - eta) * 100;
    const cost_th = E_th * P;
    const cost_real = E_real * P;
    const loss_pln_per_m3 = cost_real - cost_th;
    const loss_pln_total = loss_pln_per_m3 * V;

    const warnings: string[] = [];
    if (deltaT < 30 || deltaT > 60) warnings.push(`Nietypowe ΔT = ${deltaT.toFixed(1)} K`);
    if (eta > 1.1) warnings.push("Sprawność > 110% – sprawdź E lub m³");
    if (eta <= 0) warnings.push("Sprawność ≤ 0% – błąd danych");

    return NextResponse.json({
      ok: true,
      inputs: { period: data.period, E_DHW_GJ:E, V_DHW_m3:V, price_GJ_PLN:P, T_hot_C:Th, T_cold_C:Tc },
      thermo: { deltaT_K: deltaT, E_th_GJ_per_m3:+E_th.toFixed(6), E_real_GJ_per_m3:+E_real.toFixed(6), eta_total:+eta.toFixed(4), loss_pct:+loss_pct.toFixed(2) },
      costs: { cost_th_PLN_per_m3:+cost_th.toFixed(2), cost_real_PLN_per_m3:+cost_real.toFixed(2), loss_PLN_per_m3:+loss_pln_per_m3.toFixed(2), loss_PLN_total:+loss_pln_total.toFixed(2) },
      warnings
    });
  } catch (e:any) {
    return NextResponse.json({ ok:false, errors:[e?.message || "Server error"] }, { status:500 });
  }
}