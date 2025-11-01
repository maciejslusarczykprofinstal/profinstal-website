import { Metadata } from 'next';
import { COMPANY_INFO } from "@/lib/config/metadata";

export const metadata: Metadata = {
  title: `Kalkulator CWU - ${COMPANY_INFO.name}`,
  description: 'Kalkulator zapotrzebowania na ciepłą wodę użytkową dla budynków mieszkalnych. Oblicz parametry instalacji CWU.',
  keywords: [
    'kalkulator CWU',
    'zapotrzebowanie CWU',
    'ciepła woda użytkowa',
    'instalacja CWU',
    'obliczenia hydrauliczne',
    'budynek mieszkalny'
  ],
};

export default function CwuCalculatorLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return children;
}