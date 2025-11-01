/**
 * Stałe fizyczne dla obliczeń audytu CWU
 */
export const PHYSICS_CONSTANTS = {
  /** Ciepło właściwe wody w GJ/(m³·K) */
  WATER_SPECIFIC_HEAT: 0.004186,
  
  /** Konwersja kJ do GJ */
  KJ_TO_GJ: 1e-6,
  
  /** Gęstość wody w kg/m³ przy 15°C */
  WATER_DENSITY: 1000,
  
  /** Ciepło właściwe wody w kJ/(kg·K) */
  WATER_SPECIFIC_HEAT_KJ: 4.186,
} as const;

/**
 * Domyślne wartości temperatur
 */
export const DEFAULT_TEMPERATURES = {
  /** Domyślna temperatura ciepłej wody w °C */
  HOT_WATER: 55,
  
  /** Domyślna temperatura zimnej wody w °C */
  COLD_WATER: 8,
  
  /** Minimalna różnica temperatur w K */
  MIN_DELTA_T: 30,
  
  /** Maksymalna różnica temperatur w K */
  MAX_DELTA_T: 60,
} as const;

/**
 * Limity walidacji
 */
export const VALIDATION_LIMITS = {
  /** Maksymalna sprawność (110%) */
  MAX_EFFICIENCY: 1.1,
  
  /** Minimalna sprawność (0%) */
  MIN_EFFICIENCY: 0,
  
  /** Minimalne wartości */
  MIN_VALUES: {
    ENERGY: 0,
    VOLUME: 0,
    PRICE: 0,
  },
} as const;

/**
 * Konfiguracja formatowania
 */
export const FORMAT_CONFIG = {
  /** Miejsca po przecinku dla różnych typów wartości */
  DECIMAL_PLACES: {
    ENERGY: 6,
    EFFICIENCY: 4,
    PERCENTAGE: 2,
    CURRENCY: 2,
    TEMPERATURE: 1,
  },
} as const;