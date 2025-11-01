import { AuditResults, ApiResponse, AuditFormData, DatePeriod } from '@/lib/types';

/**
 * Wysyła zapytanie do API audytu CWU
 */
export async function submitAuditData(auditData: Record<string, unknown>): Promise<ApiResponse<AuditResults>> {
  try {
    const response = await fetch('/api/audit-cwu', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(auditData),
    });

    const result = await response.json();
    
    if (!response.ok) {
      return {
        ok: false,
        errors: result.errors || ['Błąd serwera']
      };
    }

    return {
      ok: true,
      data: result,
      warnings: result.warnings || []
    };
  } catch {
    return {
      ok: false,
      errors: ['Błąd połączenia z serwerem']
    };
  }
}

/**
 * Tworzy obiekt danych do wysłania do API
 */
export function createApiPayload(
  form: AuditFormData,
  period: DatePeriod
) {
  return {
    period,
    E_DHW_GJ: Number(form.E_DHW_GJ),
    V_DHW_m3: Number(form.V_DHW_m3),
    price_GJ_PLN: Number(form.price_GJ_PLN),
    T_hot_C: Number(form.T_hot_C),
    T_cold_C: Number(form.T_cold_C),
  };
}