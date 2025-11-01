/**
 * Dane wejściowe dla audytu CWU
 */
export interface AuditInputs {
  /** Energia ciepłej wody użytkowej w GJ */
  E_DHW_GJ: number;
  /** Objętość ciepłej wody użytkowej w m³ */
  V_DHW_m3: number;
  /** Cena energii w PLN/GJ */
  price_GJ_PLN: number;
  /** Temperatura ciepłej wody w °C */
  T_hot_C: number;
  /** Temperatura zimnej wody w °C */
  T_cold_C: number;
  /** Okres analizy */
  period: DatePeriod;
}

/**
 * Okres dat
 */
export interface DatePeriod {
  /** Data początkowa (YYYY-MM-DD) */
  from: string;
  /** Data końcowa (YYYY-MM-DD) */
  to: string;
}

/**
 * Dane formularza (przed konwersją na liczby)
 */
export interface AuditFormData {
  E_DHW_GJ: string;
  V_DHW_m3: string;
  price_GJ_PLN: string;
  T_hot_C: string;
  T_cold_C: string;
}

/**
 * Wyniki obliczeń termodynamicznych
 */
export interface ThermoResults {
  /** Różnica temperatur w Kelwinach */
  deltaT_K: number;
  /** Energia teoretyczna w GJ/m³ */
  E_th_GJ_per_m3: number;
  /** Energia rzeczywista w GJ/m³ */
  E_real_GJ_per_m3: number;
  /** Sprawność całkowita */
  eta_total: number;
  /** Straty procentowe */
  loss_pct: number;
}

/**
 * Wyniki obliczeń kosztów
 */
export interface CostResults {
  /** Koszt teoretyczny w PLN/m³ */
  cost_th_PLN_per_m3: number;
  /** Koszt rzeczywisty w PLN/m³ */
  cost_real_PLN_per_m3: number;
  /** Straty w PLN/m³ */
  loss_PLN_per_m3: number;
  /** Straty całkowite w PLN */
  loss_PLN_total: number;
}

/**
 * Pełne wyniki audytu CWU
 */
export interface AuditResults {
  /** Dane wejściowe */
  inputs: AuditInputs;
  /** Wyniki termodynamiczne */
  thermo: ThermoResults;
  /** Wyniki kosztowe */
  costs: CostResults;
  /** Ostrzeżenia */
  warnings: string[];
}

/**
 * Uproszczone wyniki dla UI (legacy)
 */
export interface SimpleAuditResult {
  loss_pct: number;
  cost_th_PLN_per_m3: number;
  cost_real_PLN_per_m3: number;
  loss_PLN_per_m3: number;
  loss_PLN_total: number;
}

/**
 * Wynik obliczeń mocy CWU
 */
export interface CwuPowerResult {
  /** Moc podstawowa w kW */
  mocPodstawowa: number;
  /** Moc ze stratami cyrkulacji w kW */
  mocZamowiona: number;
  /** Straty cyrkulacji w kW */
  stratyCyrkulacji: number;
  /** Procent strat */
  procentStrat: number;
}