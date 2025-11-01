/**
 * Formatuje liczbę do określonej liczby miejsc po przecinku
 */
export function formatNumber(value: number, decimals: number = 2): string {
  return value.toFixed(decimals);
}

/**
 * Formatuje wartość procentową
 */
export function formatPercentage(value: number, decimals: number = 1): string {
  return `${value.toFixed(decimals)}%`;
}

/**
 * Formatuje wartość finansową w PLN
 */
export function formatCurrency(value: number, decimals: number = 2): string {
  return `${value.toFixed(decimals)} zł`;
}

/**
 * Formatuje wartość z jednostką
 */
export function formatWithUnit(value: number, unit: string, decimals: number = 2): string {
  return `${value.toFixed(decimals)} ${unit}`;
}

/**
 * Formatuje zakres wartości (np. "10.5 zł/m³ • 1250.00 zł / okres")
 */
export function formatRange(value1: number, unit1: string, value2: number, unit2: string): string {
  return `${formatWithUnit(value1, unit1)} • ${formatWithUnit(value2, unit2)}`;
}