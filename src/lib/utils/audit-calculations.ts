import { AuditInputs, AuditResults } from '@/lib/types';
import { PHYSICS_CONSTANTS, DEFAULT_TEMPERATURES, VALIDATION_LIMITS } from '@/lib/config';

export function calculateAuditResults(inputs: AuditInputs): AuditResults {
  const { E_DHW_GJ: E, V_DHW_m3: V, price_GJ_PLN: P, T_hot_C: Th, T_cold_C: Tc } = inputs;

  // Obliczenia termodynamiczne
  const deltaT = Th - Tc;
  const E_th = PHYSICS_CONSTANTS.WATER_SPECIFIC_HEAT * deltaT;   // GJ/m³
  const E_real = E / V;                        // GJ/m³
  const eta = E_th / E_real;
  const loss_pct = (1 - eta) * 100;

  // Obliczenia kosztów
  const cost_th = E_th * P;
  const cost_real = E_real * P;
  const loss_pln_per_m3 = cost_real - cost_th;
  const loss_pln_total = loss_pln_per_m3 * V;

  // Generowanie ostrzeżeń
  const warnings = generateWarnings(deltaT, eta);

  return {
    inputs,
    thermo: {
      deltaT_K: deltaT,
      E_th_GJ_per_m3: +E_th.toFixed(6),
      E_real_GJ_per_m3: +E_real.toFixed(6),
      eta_total: +eta.toFixed(4),
      loss_pct: +loss_pct.toFixed(2)
    },
    costs: {
      cost_th_PLN_per_m3: +cost_th.toFixed(2),
      cost_real_PLN_per_m3: +cost_real.toFixed(2),
      loss_PLN_per_m3: +loss_pln_per_m3.toFixed(2),
      loss_PLN_total: +loss_pln_total.toFixed(2)
    },
    warnings
  };
}

function generateWarnings(deltaT: number, eta: number): string[] {
  const warnings: string[] = [];

  if (deltaT < DEFAULT_TEMPERATURES.MIN_DELTA_T || deltaT > DEFAULT_TEMPERATURES.MAX_DELTA_T) {
    warnings.push(`Nietypowe ΔT = ${deltaT.toFixed(1)} K`);
  }

  if (eta > VALIDATION_LIMITS.MAX_EFFICIENCY) {
    warnings.push("Sprawność > 110% – sprawdź E lub m³");
  }

  if (eta <= VALIDATION_LIMITS.MIN_EFFICIENCY) {
    warnings.push("Sprawność ≤ 0% – błąd danych");
  }

  return warnings;
}