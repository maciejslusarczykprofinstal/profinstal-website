"use client";

import { useState } from "react";
import AuditForm from "@/components/forms/AuditForm";
import WarningsAlert from "@/components/ui/WarningsAlert";
import ResultsDisplay from "@/components/ui/ResultsDisplay";
import Breadcrumbs from "@/components/ui/Breadcrumbs";
import { submitAuditData, createApiPayload } from "@/lib/utils/api";
import { SimpleAuditResult, AuditFormData, DatePeriod } from "@/lib/types";
import { DEFAULT_TEMPERATURES } from "@/lib/config";

export default function AudytCWU() {
  const [form, setForm] = useState<AuditFormData>({
    E_DHW_GJ: "",
    V_DHW_m3: "",
    price_GJ_PLN: "",
    T_hot_C: DEFAULT_TEMPERATURES.HOT_WATER.toString(),
    T_cold_C: DEFAULT_TEMPERATURES.COLD_WATER.toString(),
  });
  const [period, setPeriod] = useState<DatePeriod>({ from: "", to: "" });
  const [res, setRes] = useState<SimpleAuditResult | null>(null);
  const [warn, setWarn] = useState<string[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  const onChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setForm({ ...form, [e.target.name]: e.target.value });
  };

  async function calc() {
    setIsLoading(true);
    setRes(null);
    setWarn([]);

    try {
      const payload = createApiPayload(form, period);
      const response = await submitAuditData(payload);

      if (!response.ok) {
        setWarn(response.errors || ["Błąd danych"]);
        return;
      }

      if (response.data) {
        setWarn(response.warnings || []);
        setRes({
          loss_pct: response.data.thermo.loss_pct,
          cost_th_PLN_per_m3: response.data.costs.cost_th_PLN_per_m3,
          cost_real_PLN_per_m3: response.data.costs.cost_real_PLN_per_m3,
          loss_PLN_per_m3: response.data.costs.loss_PLN_per_m3,
          loss_PLN_total: response.data.costs.loss_PLN_total,
        });
      }
    } catch {
      setWarn(["Wystąpił nieoczekiwany błąd"]);
    } finally {
      setIsLoading(false);
    }
  }

  return (
    <main className="mx-auto max-w-4xl px-6 py-12 space-y-6">
      <Breadcrumbs items={[
        { label: 'Audyt CWU' }
      ]} />
      
      <h1 className="text-3xl font-semibold">Audyt CWU</h1>

      <AuditForm 
        form={form}
        period={period}
        onChange={onChange}
        setPeriod={setPeriod}
        onSubmit={calc}
        isLoading={isLoading}
      />

      <WarningsAlert warnings={warn} />

      {res && <ResultsDisplay results={res} />}
    </main>
  );
}