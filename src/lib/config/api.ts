/**
 * Endpointy API
 */
export const API_ENDPOINTS = {
  /** Audyt CWU */
  AUDIT_CWU: '/api/audit-cwu',
} as const;

/**
 * Konfiguracja żądań HTTP
 */
export const HTTP_CONFIG = {
  /** Timeout dla żądań w ms */
  TIMEOUT: 30000,
  
  /** Retry attempts */
  MAX_RETRIES: 3,
  
  /** Headers */
  DEFAULT_HEADERS: {
    'Content-Type': 'application/json',
  },
} as const;

/**
 * Kody statusów HTTP
 */
export const HTTP_STATUS = {
  OK: 200,
  BAD_REQUEST: 400,
  UNAUTHORIZED: 401,
  FORBIDDEN: 403,
  NOT_FOUND: 404,
  INTERNAL_SERVER_ERROR: 500,
} as const;

/**
 * Komunikaty błędów
 */
export const ERROR_MESSAGES = {
  NETWORK_ERROR: 'Błąd połączenia z serwerem',
  SERVER_ERROR: 'Błąd serwera',
  VALIDATION_ERROR: 'Błąd walidacji danych',
  TIMEOUT_ERROR: 'Przekroczono czas oczekiwania',
  UNKNOWN_ERROR: 'Wystąpił nieoczekiwany błąd',
} as const;