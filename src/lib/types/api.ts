/**
 * Generyczna odpowiedź API
 */
export interface ApiResponse<T = unknown> {
  /** Status operacji */
  ok: boolean;
  /** Dane odpowiedzi (jeśli sukces) */
  data?: T;
  /** Lista błędów (jeśli niepowodzenie) */
  errors?: string[];
  /** Lista ostrzeżeń */
  warnings?: string[];
}

/**
 * Wynik walidacji
 */
export interface ValidationResult {
  /** Czy dane są poprawne */
  isValid: boolean;
  /** Lista błędów walidacji */
  errors: string[];
}

/**
 * Status operacji
 */
export type OperationStatus = 'idle' | 'loading' | 'success' | 'error';

/**
 * Parametry błędu API
 */
export interface ApiError {
  /** Kod błędu HTTP */
  status?: number;
  /** Wiadomość błędu */
  message: string;
  /** Szczegóły błędu */
  details?: Record<string, unknown>;
}