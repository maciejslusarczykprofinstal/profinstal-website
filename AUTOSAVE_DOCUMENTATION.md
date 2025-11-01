# Funkcjonalność Autosave - Kalkulator CWU

## Opis funkcji

Kalkulator CWU automatycznie zapisuje wprowadzone dane w localStorage przeglądarki i przywraca je przy ponownym wejściu na stronę.

## Główne funkcjonalności

### 🔄 **Automatyczne zapisywanie**
- Dane są zapisywane po każdej zmianie w formularzu
- Nie wymaga żadnej akcji ze strony użytkownika
- Zapisuje się w localStorage pod kluczem `cwu-calculator-data`

### 📥 **Automatyczne wczytywanie**
- Przy wejściu na stronę dane są automatycznie przywracane
- Jeśli są zapisane dane, wyświetla się informacyjny komunikat
- Domyślne wartości są zachowane jeśli brak zapisanych danych

### 🗑️ **Czyszczenie danych**
- Przycisk "wyczyść" w prawym górnym rogu kalkulatora
- Usuwa zapisane dane z localStorage
- Przywraca domyślne wartości formularza

### 📤 **Eksport danych**
- Możliwość eksportu danych do pliku JSON
- Zawiera dane formularza, datę eksportu i wersję
- Plik można zapisać lokalnie jako backup

### 📥 **Import danych**
- Możliwość importu danych z pliku JSON
- Automatyczna walidacja formatu pliku
- Resetuje wyniki obliczeń przy imporcie nowych danych

## Interfejs użytkownika

### Wskaźniki stanu
```
[✓] Dane zapisywane automatycznie    [🗑️]
```

### Komunikat o wczytanych danych
```
ℹ️ Wczytano poprzednio zapisane dane. Możesz kontynuować obliczenia lub wyczyścić dane powyżej.
```

### Dodatkowe opcje
```
[📥 Eksportuj dane]  [📤 Importuj dane]
```

## Struktura zapisywanych danych

```json
{
  "liczba_mieszkan": "50",
  "liczba_pionow": "4", 
  "temp_zimnej_wody": "8",
  "temp_cwu": "55",
  "procent_strat_cyrkulacji": "10"
}
```

## Struktura eksportowanych danych

```json
{
  "formData": {
    "liczba_mieszkan": "50",
    "liczba_pionow": "4",
    "temp_zimnej_wody": "8", 
    "temp_cwu": "55",
    "procent_strat_cyrkulacji": "10"
  },
  "exportDate": "2025-11-01T...",
  "version": "1.0"
}
```

## Obsługa błędów

- **localStorage niedostępny**: Graceful degradation, aplikacja działa normalnie
- **Błędne dane w localStorage**: Automatyczne czyszczenie i powrót do domyślnych wartości
- **Błąd importu**: Wyświetlenie komunikatu błędu, zachowanie obecnych danych

## Korzyści UX

### 🎯 **Wygoda użytkownika**
- Nie traci wprowadzonych danych przy przypadkowym odświeżeniu
- Może kontynuować pracę w dowolnym momencie
- Szybkie przywracanie często używanych konfiguracji

### 💾 **Backup i sharing**
- Możliwość zapisania konfiguracji do pliku
- Udostępnianie konfiguracji między urządzeniami
- Archiwizacja różnych wariantów obliczeń

### 🔒 **Prywatność**
- Dane przechowywane lokalnie w przeglądarce
- Brak transmisji danych do serwera (dla autosave)
- Pełna kontrola nad danymi przez użytkownika

## Implementacja techniczna

- **React useState + useEffect** dla zarządzania stanem
- **localStorage API** dla trwałości danych
- **File API** dla eksportu/importu
- **JSON serialization** dla formatu danych
- **Error boundaries** dla obsługi błędów