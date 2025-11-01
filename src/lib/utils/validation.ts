import { AuditInputs, ValidationResult } from '@/lib/types';
import { DEFAULT_TEMPERATURES, VALIDATION_LIMITS } from '@/lib/config';

export function validateAuditInputs(data: Record<string, unknown>): ValidationResult {
  const errors: string[] = [];

  // Sprawdź wymagane pola
  const requiredFields = ["E_DHW_GJ", "V_DHW_m3", "price_GJ_PLN", "T_hot_C", "T_cold_C", "period"];
  for (const field of requiredFields) {
    if (data[field] === undefined) {
      errors.push(`${field} jest wymagane`);
    }
  }

  if (errors.length > 0) {
    return { isValid: false, errors };
  }

  // Walidacja wartości liczbowych
  const E = Number(data.E_DHW_GJ);
  const V = Number(data.V_DHW_m3);
  const P = Number(data.price_GJ_PLN);
  const Th = Number(data.T_hot_C ?? DEFAULT_TEMPERATURES.HOT_WATER);
  const Tc = Number(data.T_cold_C ?? DEFAULT_TEMPERATURES.COLD_WATER);

  if (!(E > VALIDATION_LIMITS.MIN_VALUES.ENERGY)) errors.push("E_DHW_GJ > 0");
  if (!(V > VALIDATION_LIMITS.MIN_VALUES.VOLUME)) errors.push("V_DHW_m3 > 0");
  if (!(P > VALIDATION_LIMITS.MIN_VALUES.PRICE)) errors.push("price_GJ_PLN > 0");
  if (!(Th > Tc)) errors.push("T_hot_C > T_cold_C");

  return {
    isValid: errors.length === 0,
    errors
  };
}

export function parseAuditInputs(data: Record<string, unknown>): AuditInputs {
  return {
    E_DHW_GJ: Number(data.E_DHW_GJ),
    V_DHW_m3: Number(data.V_DHW_m3),
    price_GJ_PLN: Number(data.price_GJ_PLN),
    T_hot_C: Number(data.T_hot_C ?? DEFAULT_TEMPERATURES.HOT_WATER),
    T_cold_C: Number(data.T_cold_C ?? DEFAULT_TEMPERATURES.COLD_WATER),
    period: data.period as { from: string; to: string }
  };
}