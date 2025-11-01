import { Metadata } from 'next';

/**
 * Podstawowe informacje o firmie
 */
export const COMPANY_INFO = {
  name: 'PROF INSTAL',
  email: 'kontakt@profinstal.info',
  description: 'HVAC | CWU | Audyty energetyczne dla spółdzielni i wspólnot',
  tagline: 'Obliczenia mocy, analiza strat cyrkulacji, raporty DOCX, doradztwo modernizacyjne.',
} as const;

/**
 * Metadata dla strony głównej
 */
export const HOME_METADATA: Metadata = {
  title: `${COMPANY_INFO.name} - ${COMPANY_INFO.description}`,
  description: `${COMPANY_INFO.description}. ${COMPANY_INFO.tagline}`,
  keywords: [
    'HVAC',
    'CWU',
    'audyt energetyczny',
    'spółdzielnia mieszkaniowa',
    'wspólnota mieszkaniowa',
    'straty cyrkulacji',
    'obliczenia mocy',
    'modernizacja',
    'efektywność energetyczna'
  ],
  authors: [{ name: COMPANY_INFO.name }],
  creator: COMPANY_INFO.name,
  publisher: COMPANY_INFO.name,
  openGraph: {
    title: `${COMPANY_INFO.name} - ${COMPANY_INFO.description}`,
    description: COMPANY_INFO.tagline,
    type: 'website',
    locale: 'pl_PL',
    siteName: COMPANY_INFO.name,
  },
  twitter: {
    card: 'summary_large_image',
    title: `${COMPANY_INFO.name} - ${COMPANY_INFO.description}`,
    description: COMPANY_INFO.tagline,
  },
  robots: {
    index: true,
    follow: true,
  },
};

/**
 * Metadata dla strony audytu CWU
 */
export const AUDIT_CWU_METADATA: Metadata = {
  title: `Audyt CWU - ${COMPANY_INFO.name}`,
  description: 'Kalkulator strat cyrkulacji ciepłej wody użytkowej. Oblicz straty energetyczne i finansowe w instalacji CWU.',
  keywords: [
    'audyt CWU',
    'cyrkulacja CWU',
    'straty energetyczne',
    'ciepła woda użytkowa',
    'kalkulator strat',
    'efektywność energetyczna',
    'optymalizacja CWU'
  ],
  openGraph: {
    title: `Audyt CWU - ${COMPANY_INFO.name}`,
    description: 'Kalkulator strat cyrkulacji ciepłej wody użytkowej',
    type: 'website',
  },
};