import { AuditFormData, DatePeriod, SimpleAuditResult, CwuPowerResult } from './audit';

/**
 * Dane formularza kalkulatora CWU
 */
export interface CwuCalculatorData {
  /** Liczba mieszkań */
  liczba_mieszkan: string;
  /** Liczba pionów instalacji */
  liczba_pionow: string;
  /** Temperatura zimnej wody w °C */
  temp_zimnej_wody: string;
  /** Temperatura CWU w °C */
  temp_cwu: string;
  /** Procent strat na cyrkulacji */
  procent_strat_cyrkulacji: string;
}

/**
 * Props komponentu kalkulatora CWU
 */
export interface CwuCalculatorProps {
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
  /** Handler wyniku obliczeń */
  onCalculate?: (data: CwuCalculatorData) => void;
}

/**
 * Props komponentu wyników CWU
 */
export interface CwuResultsDisplayProps {
  /** Wyniki obliczeń do wyświetlenia */
  results: CwuPowerResult;
}

/**
 * Props komponentu formularza audytu
 */
export interface AuditFormProps {
  /** Dane formularza */
  form: AuditFormData;
  /** Okres analizy */
  period: DatePeriod;
  /** Handler zmiany pól formularza */
  onChange: (e: React.ChangeEvent<HTMLInputElement>) => void;
  /** Setter okresu */
  setPeriod: (period: DatePeriod) => void;
  /** Handler submitu formularza */
  onSubmit: () => void;
  /** Stan ładowania */
  isLoading?: boolean;
}

/**
 * Props komponentu ostrzeżeń
 */
export interface WarningsAlertProps {
  /** Lista ostrzeżeń do wyświetlenia */
  warnings: string[];
}

/**
 * Props komponentu wyników
 */
export interface ResultsDisplayProps {
  /** Wyniki audytu do wyświetlenia */
  results: SimpleAuditResult;
}

/**
 * Props komponentu nagłówka
 */
export interface HeaderProps {
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
}

/**
 * Props komponentu sekcji hero
 */
export interface HeroSectionProps {
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
}

/**
 * Props komponentu siatki usług
 */
export interface ServicesGridProps {
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
}

/**
 * Props komponentu Card
 */
export interface CardProps {
  /** Zawartość karty */
  children: React.ReactNode;
  /** Opcjonalny link */
  href?: string;
  /** Klasy CSS */
  className?: string;
  /** Indeks animacji */
  index?: number;
}

/**
 * Props komponentu ServiceCard
 */
export interface ServiceCardProps {
  /** Tytuł usługi */
  title: string;
  /** Opis usługi */
  description: string;
  /** Opcjonalny link */
  href?: string;
  /** Indeks dla animacji */
  index?: number;
}

/**
 * Props komponentu kontaktu
 */
export interface ContactSectionProps {
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
}

/**
 * Dane wiadomości AI Assistant
 */
export interface AiMessage {
  /** Unikalny identyfikator wiadomości */
  id: string;
  /** Rola nadawcy: user lub assistant */
  role: 'user' | 'assistant';
  /** Treść wiadomości */
  content: string;
  /** Timestamp utworzenia */
  timestamp: Date;
}

/**
 * Props komponentu AI Assistant
 */
export interface AiAssistantProps {
  /** Wyniki obliczeń CWU do kontekstu */
  calculationResults?: CwuPowerResult;
  /** Dane wejściowe kalkulatora do kontekstu */
  inputData?: CwuCalculatorData;
  /** Opcjonalne dodatkowe klasy CSS */
  className?: string;
}