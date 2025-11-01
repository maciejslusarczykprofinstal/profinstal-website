// Eksport wszystkich konfiguracji
export * from './constants';
export * from './metadata';
export * from './navigation';
export * from './api';

/**
 * Środowisko aplikacji
 */
export const APP_ENV = {
  isDevelopment: process.env.NODE_ENV === 'development',
  isProduction: process.env.NODE_ENV === 'production',
  isTest: process.env.NODE_ENV === 'test',
} as const;

/**
 * Konfiguracja aplikacji
 */
export const APP_CONFIG = {
  /** Nazwa aplikacji */
  name: 'PROF INSTAL',
  
  /** Wersja aplikacji */
  version: '1.0.0',
  
  /** Domyślny język */
  defaultLocale: 'pl',
  
  /** Obsługiwane języki */
  supportedLocales: ['pl'] as const,
} as const;