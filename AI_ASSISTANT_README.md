# Konfiguracja AI Assistant

## Wymagania

Aby moduł AI Assistant działał poprawnie, potrzebujesz:

1. **Klucz API OpenAI**
   - Załóż konto na https://platform.openai.com
   - Wygeneruj API key w sekcji API keys
   - Skopiuj klucz do pliku `.env.local`

## Instalacja

1. **Utwórz plik `.env.local`** w głównym katalogu projektu:
```bash
OPENAI_API_KEY=sk-your-actual-openai-api-key-here
```

2. **Restart serwera deweloperskiego**:
```bash
npm run dev
```

## Funkcjonalności

### 🤖 AI Ekspert CWU
- **Lokalizacja**: Floating button w prawym dolnym rogu ekranu
- **Dostępność**: Na wszystkich stronach aplikacji
- **Kontekst**: Automatycznie przekazuje dane z kalkulatora CWU

### 📊 Integracja z kalkulatorem
Gdy użytkownik wykonuje obliczenia CWU, AI otrzymuje pełny kontekst:
- Dane wejściowe (liczba mieszkań, temperatury, etc.)
- Wyniki obliczeń (moc podstawowa, zamówiona, straty)
- Możliwość analizy i rekomendacji

### 💬 Przykładowe pytania
- "Czy moc 25 kW wystarczy dla 50 mieszkań?"
- "Jak zmniejszyć straty cyrkulacji?"
- "Jaki kocioł polecasz dla tych parametrów?"
- "Czy warto modernizować stary system?"

## API Endpoints

### POST `/api/assistant`
```json
{
  "message": "Pytanie użytkownika",
  "context": {
    "calculationResults": {...},
    "inputData": {...}
  }
}
```

**Odpowiedź**:
```json
{
  "response": "Odpowiedź AI eksperta",
  "timestamp": "2025-11-01T..."
}
```

## Ograniczenia bez API key

Jeśli nie skonfigurujesz `OPENAI_API_KEY`:
- Komponent AI Assistant będzie się wyświetlał
- API zwróci błąd konfiguracji
- Aplikacja będzie działała normalnie (graceful degradation)

## Bezpieczeństwo

- Klucz API jest używany tylko po stronie serwera
- Nie jest eksponowany w frontend bundlu
- Komunikacja odbywa się przez bezpieczny endpoint `/api/assistant`

## Koszty

- Używamy modelu `gpt-3.5-turbo` (ekonomiczny)
- Limit 500 tokenów na odpowiedź
- Niska temperatura (0.1) dla spójnych odpowiedzi technicznych