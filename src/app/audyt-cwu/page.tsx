"use client";

import { useState } from "react";

type Result = {
  loss_pct: number;
  cost_th_PLN_per_m3: number;
  cost_real_PLN_per_m3: number;
  loss_PLN_per_m3: number;
  loss_PLN_total: number;
};

export default function AudytCWU() {
  const [form, setForm] = useState({
    E_DHW_GJ: "",
    V_DHW_m3: "",
    price_GJ_PLN: "",
    T_hot_C: "55",
    T_cold_C: "8",
  });
  const [period, setPeriod] = useState({ from: "", to: "" });
  const [res, setRes] = useState<Result | null>(null);
  const [warn, setWarn] = useState<string[]>([]);

  const onChange = (e: any) => setForm({ ...form, [e.target.name]: e.target.value });

  async function calc() {
    setRes(null);
    setWarn([]);
    const body = {
      period,
      E_DHW_GJ: Number(form.E_DHW_GJ),
      V_DHW_m3: Number(form.V_DHW_m3),
      price_GJ_PLN: Number(form.price_GJ_PLN),
      T_hot_C: Number(form.T_hot_C),
      T_cold_C: Number(form.T_cold_C),
    };
    const r = await fetch("/api/audit-cwu", { method: "POST", body: JSON.stringify(body) });
    const j = await r.json();
    if (!j.ok) { setWarn(j.errors || ["Błąd danych"]); return; }
    setWarn(j.warnings || []);
    setRes({
      loss_pct: j.thermo.loss_pct,
      cost_th_PLN_per_m3: j.costs.cost_th_PLN_per_m3,
      cost_real_PLN_per_m3: j.costs.cost_real_PLN_per_m3,
      loss_PLN_per_m3: j.costs.loss_PLN_per_m3,
      loss_PLN_total: j.costs.loss_PLN_total,
    });
  }

  return (
    <main className="mx-auto max-w-4xl px-6 py-12 space-y-6">
      <h1 className="text-3xl font-semibold">Audyt CWU</h1>

      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        <input className="border rounded-lg p-3" type="date" value={period.from} onChange={e=>setPeriod({...period,from:e.target.value})} />
        <input className="border rounded-lg p-3" type="date" value={period.to} onChange={e=>setPeriod({...period,to:e.target.value})} />
        <input className="border rounded-lg p-3" placeholder="Cena zł/GJ" name="price_GJ_PLN" onChange={onChange}/>
        <input className="border rounded-lg p-3" placeholder="E_DHW_GJ" name="E_DHW_GJ" onChange={onChange}/>
        <input className="border rounded-lg p-3" placeholder="V_DHW_m3" name="V_DHW_m3" onChange={onChange}/>
        <input className="border rounded-lg p-3" placeholder="T_hot °C" name="T_hot_C" defaultValue={form.T_hot_C} onChange={onChange}/>
        <input className="border rounded-lg p-3" placeholder="T_cold °C" name="T_cold_C" defaultValue={form.T_cold_C} onChange={onChange}/>
        <button onClick={calc} className="px-5 py-3 rounded-lg bg-black text-white md:col-span-3">OBLICZ STRATY</button>
      </div>

      {warn.length > 0 && (
        <div className="p-4 rounded-lg bg-yellow-50 border text-sm">
          {warn.map((w,i)=><div key={i}>• {w}</div>)}
        </div>
      )}

      {res && (
        <div className="grid md:grid-cols-2 gap-4">
          <div className="p-6 border rounded-2xl">
            <div className="text-sm text-gray-500">Straty całkowite</div>
            <div className="text-3xl font-semibold">{res.loss_pct.toFixed(1)}%</div>
            <div className="text-sm">{res.loss_PLN_per_m3.toFixed(2)} zł/m³ • {res.loss_PLN_total.toFixed(2)} zł / okres</div>
          </div>
          <div className="p-6 border rounded-2xl">
            <div className="text-sm text-gray-500">Koszty</div>
            <div className="text-sm">Teoria: {res.cost_th_PLN_per_m3.toFixed(2)} zł/m³</div>
            <div className="text-sm">Rzeczywiste: {res.cost_real_PLN_per_m3.toFixed(2)} zł/m³</div>
          </div>
        </div>
      )}
    </main>
  );
}