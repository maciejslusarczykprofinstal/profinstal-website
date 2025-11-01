/**
 * Element nawigacji
 */
export interface NavigationItem {
  /** Tekst wyświetlany */
  label: string;
  /** URL lub hash */
  href: string;
  /** Czy link otwiera się w nowej karcie */
  external?: boolean;
}

/**
 * Główna nawigacja strony
 */
export const MAIN_NAVIGATION: NavigationItem[] = [
  {
    label: 'Home',
    href: '/',
  },
  {
    label: 'Usługi',
    href: '#uslugi',
  },
  {
    label: 'Kontakt',
    href: '#kontakt',
  },
  {
    label: 'Audyt CWU',
    href: '/audyt-cwu',
  },
  {
    label: 'Kalkulator CWU',
    href: '/kalkulator-cwu',
  },
  {
    label: 'AI Ekspert',
    href: '/ai-assistant',
  },
] as const;

/**
 * Konfiguracja usług na stronie głównej
 */
export interface ServiceCard {
  /** Tytuł usługi */
  title: string;
  /** Opis usługi */
  description: string;
  /** Opcjonalny link */
  href?: string;
}

export const SERVICES: ServiceCard[] = [
  {
    title: 'Audyt CWU',
    description: 'Straty [%], zł/m³ i rekomendacje na podstawie liczników.',
    href: '/audyt-cwu',
  },
  {
    title: 'Kalkulator CWU',
    description: 'Obliczenia mocy kotła i instalacji CWU.',
    href: '/kalkulator-cwu',
  },
  {
    title: 'AI Ekspert CWU',
    description: 'Konsultacje AI w zakresie modernizacji i doboru urządzeń.',
    href: '/ai-assistant',
  },
  {
    title: 'HVAC/Serwis',
    description: 'Klimatyzacja, wentylacja, równoważenie instalacji.',
  },
  {
    title: 'Raporty DOCX',
    description: 'Wnioski techniczne i finansowe dla zarządów.',
  },
] as const;

/**
 * Przyciski CTA (Call To Action)
 */
export interface CTAButton {
  /** Tekst przycisku */
  label: string;
  /** Link */
  href: string;
  /** Styl przycisku */
  variant: 'primary' | 'secondary';
}

export const CTA_BUTTONS: CTAButton[] = [
  {
    label: 'Uruchom Audyt CWU',
    href: '/audyt-cwu',
    variant: 'primary',
  },
  {
    label: 'Kontakt',
    href: '#kontakt',
    variant: 'secondary',
  },
] as const;