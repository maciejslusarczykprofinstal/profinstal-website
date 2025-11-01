# Routing i Nawigacja - PROF INSTAL

## Mapa routingu

### 🏠 **Główne strony**
```
/              → Homepage (strona główna)
/audyt-cwu     → Audyt CWU (analiza strat)
/kalkulator-cwu → Kalkulator CWU (obliczenia mocy)
/ai-assistant  → AI Ekspert CWU (konsultacje)
```

### 🔗 **API Endpoints**
```
/api/audit-cwu    → POST - obliczenia audytu + DOCX
/api/assistant    → POST - konsultacje AI OpenAI
```

## Komponenty nawigacji

### 📋 **Header Navigation**
- **Lokalizacja**: `src/components/layout/Header.tsx`
- **Funkcjonalność**: 
  - Logo z linkiem do strony głównej
  - Menu nawigacyjne z active states
  - Obsługa hash linków (#uslugi, #kontakt)
  - Next.js Link dla lepszej wydajności

### 🔗 **Active Link Component**
- **Lokalizacja**: `src/components/ui/ActiveLink.tsx`
- **Funkcjonalność**:
  - Automatyczne oznaczanie aktywnej strony
  - `usePathname()` hook dla wykrywania lokalizacji
  - Customizowalne style dla active state

### 🍞 **Breadcrumbs Component**
- **Lokalizacja**: `src/components/ui/Breadcrumbs.tsx`
- **Funkcjonalność**:
  - Ścieżka nawigacyjna na podstronach
  - Link do Home + aktualna lokalizacja
  - Ikony separatorów

## Struktura nawigacji

### **Menu główne** (Header)
```
[PROF INSTAL] → Home | Usługi | Kontakt | Audyt CWU | Kalkulator CWU | AI Ekspert
```

### **Services Grid** (Homepage)
- Klikalne karty usług z linkami
- Hover effects i visual feedback
- "Przejdź →" dla usług z linkami

### **CTA Buttons** (Homepage)
- "Uruchom Audyt CWU" → `/audyt-cwu`
- "Kontakt" → `#kontakt` (scroll to section)

## Konfiguracja routingu

### **Navigation Config**
```typescript
// src/lib/config/navigation.ts
export const MAIN_NAVIGATION: NavigationItem[] = [
  { label: 'Home', href: '/' },
  { label: 'Usługi', href: '#uslugi' },
  { label: 'Kontakt', href: '#kontakt' },
  { label: 'Audyt CWU', href: '/audyt-cwu' },
  { label: 'Kalkulator CWU', href: '/kalkulator-cwu' },
  { label: 'AI Ekspert', href: '/ai-assistant' },
];
```

### **Services Config**
```typescript
export const SERVICES: ServiceCard[] = [
  {
    title: 'Audyt CWU',
    description: 'Straty [%], zł/m³ i rekomendacje...',
    href: '/audyt-cwu', // Klikalna karta
  },
  {
    title: 'Kalkulator CWU', 
    description: 'Obliczenia mocy kotła...',
    href: '/kalkulator-cwu', // Klikalna karta
  },
  // ...
];
```

## Funkcjonalności routingu

### ✅ **Next.js App Router**
- **Server Components** dla lepszej wydajności
- **Client Components** gdzie potrzeba interaktywności
- **Automatic code splitting** per route
- **Static generation** dla stron bez dynamicznych danych

### ✅ **Link Performance**
- **Prefetching** - automatyczne ładowanie w tle
- **Client-side navigation** - brak pełnych refresh
- **Optimized bundles** - tylko potrzebny kod

### ✅ **UX Enhancements**
- **Active states** - wizualne oznaczenie aktywnej strony
- **Breadcrumbs** - ścieżka nawigacyjna na podstronach
- **Hover effects** - feedback dla interaktywnych elementów
- **Transition animations** - gładkie przejścia między stanami

### ✅ **Accessibility**
- **Semantic HTML** - nav, header, main elements
- **Keyboard navigation** - Tab support
- **ARIA labels** - dla screen readers
- **Focus management** - właściwe focus states

## Mapowanie URL → Komponent

```
/                 → src/app/page.tsx (Homepage)
                    ├── Header
                    ├── HeroSection  
                    ├── ServicesGrid
                    └── ContactSection

/audyt-cwu        → src/app/audyt-cwu/page.tsx
                    ├── Breadcrumbs
                    ├── AuditForm
                    ├── WarningsAlert  
                    └── ResultsDisplay

/kalkulator-cwu   → src/app/kalkulator-cwu/page.tsx
                    ├── Breadcrumbs
                    ├── CwuCalculator
                    └── AiAssistant (floating)

/ai-assistant     → src/app/ai-assistant/page.tsx
                    ├── Breadcrumbs
                    ├── Instrukcje/FAQ
                    └── AiAssistant (main)
```

## SEO i Metadane

### **Page Metadata**
- Każda strona ma własne `<title>` i `<meta>`
- Open Graph tags dla social sharing
- Structured data dla wyszukiwarek

### **URL Structure**
- Czytelne, SEO-friendly URLs
- Polskie nazwy w URL (audyt-cwu, kalkulator-cwu)
- Consistent naming convention

## Przyszłe rozszerzenia

### **Potencjalne nowe routes**
- `/projekty` - portfolio realizacji
- `/cennik` - tabela cen usług  
- `/blog` - artykuły techniczne
- `/kontakt` - dedykowana strona kontaktowa
- `/o-nas` - informacje o firmie

### **Funkcjonalności routing**
- **Search** - wyszukiwarka w treściach
- **Filters** - filtrowanie usług/projektów
- **Pagination** - dla blogów/projektów
- **Language switching** - jeśli potrzeba i18n