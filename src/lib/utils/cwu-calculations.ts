import type { CwuCalculatorData } from '@/lib/types';

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

/**
 * Stałe do obliczeń CWU
 */
const CWU_CONSTANTS = {
  /** Średnie zapotrzebowanie na CWU na mieszkanie w litrach/dobę */
  ZUZYCIE_NA_MIESZKANIE: 120,
  /** Ciepło właściwe wody w kJ/(kg·K) */
  CIEPLO_WLASCIWE: 4.186,
  /** Gęstość wody w kg/l */
  GESTOSC_WODY: 1,
  /** Współczynnik simultaneity (nie wszyscy używają jednocześnie) */
  WSPOLCZYNNIK_SIMULTANEITY: 0.3,
  /** Przelicznik kJ/h na kW */
  KJ_H_TO_KW: 3600,
} as const;

/**
 * Oblicza moc zamówioną dla instalacji CWU
 */
export function calculatePower(data: CwuCalculatorData): CwuPowerResult {
  // Konwersja stringów na liczby
  const liczbaMieszkan = Number(data.liczba_mieszkan) || 0;
  // const liczbaPionow = Number(data.liczba_pionow) || 1; // TODO: Użyć w przyszłych obliczeniach
  const tempZimnej = Number(data.temp_zimnej_wody) || 8;
  const tempCwu = Number(data.temp_cwu) || 55;
  const procentStrat = Number(data.procent_strat_cyrkulacji) || 0;

  // Różnica temperatur
  const deltaT = tempCwu - tempZimnej;

  // Całkowite zapotrzebowanie na CWU w litrach/dobę
  const zapotrzebowanieDobowe = liczbaMieszkan * CWU_CONSTANTS.ZUZYCIE_NA_MIESZKANIE;

  // Zapotrzebowanie godzinowe z uwzględnieniem współczynnika simultaneity
  const zapotrzebowanieGodzinowe = zapotrzebowanieDobowe * CWU_CONSTANTS.WSPOLCZYNNIK_SIMULTANEITY;

  // Moc podstawowa w kW (bez strat cyrkulacji)
  const mocPodstawowa = (
    zapotrzebowanieGodzinowe * 
    CWU_CONSTANTS.GESTOSC_WODY * 
    CWU_CONSTANTS.CIEPLO_WLASCIWE * 
    deltaT
  ) / CWU_CONSTANTS.KJ_H_TO_KW;

  // Straty cyrkulacji w kW
  const stratyCyrkulacji = (mocPodstawowa * procentStrat) / 100;

  // Moc zamówiona (z uwzględnieniem strat)
  const mocZamowiona = mocPodstawowa + stratyCyrkulacji;

  return {
    mocPodstawowa: Number(mocPodstawowa.toFixed(2)),
    mocZamowiona: Number(mocZamowiona.toFixed(2)),
    stratyCyrkulacji: Number(stratyCyrkulacji.toFixed(2)),
    procentStrat
  };
}