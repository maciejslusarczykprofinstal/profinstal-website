
# PROF INSTAL — wersja web (Flask)

Minimalny port aplikacji Tkinter → Flask + Jinja2.

## Jak uruchomić
```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# Linux/Mac: source .venv/bin/activate
pip install -r requirements.txt
python app.py
```
Aplikacja nasłuchuje na http://127.0.0.1:5000/

## Struktura
- `app.py` — trasy Flask, formularz wejściowy i widok wyników
- `calc.py` — wyodrębniona logika obliczeń (re-use w GUI/Web)
- `templates/` — szablony Jinja2 (`index.html`, `result.html`)
- `static/style.css` — proste style
- `/export/docx`, `/export/pdf` — eksport wyników (wymaga `python-docx` i `reportlab`)

## Debug w VS Code
Wybierz konfigurację „Debug Flask Web App (app.py)” w `launch.json`.

## Uwaga
To wersja startowa. Możesz rozbudować o:
- autoryzację i RODO,
- generowanie pism reklamacyjnych z danymi adresowymi,
- upload logo,
- bazę miast/taryf (z bazy danych),
- API JSON `/api/calc`.
