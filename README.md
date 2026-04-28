# Transportation Data Extraction

Two-stage pipeline that extracts structured trip data from supplier transportation invoices (PDF/Excel) and enriches each trip with real driving distance & duration via Google Maps — turning unstructured monthly invoice batches into a clean, analyzable dataset.

Replaces a manual review where finance had to read each invoice page-by-page, type out trip details into Excel, and look up routes on Google Maps one by one.

## What it does

### Stage 1 — `תיקון לנושא סהכ חשבוניות.py` (Invoice extraction via Gemini)
- Iterates over invoice files in `חשבוניות/` (PDFs and Excels).
- Sends each file to **Google Gemini** with a structured extraction prompt (date, time, origin, waypoints, destination, quantity, prices incl/excl VAT).
- Falls back across model versions (`gemini-2.5-flash` → `gemini-2.0-flash` → `gemini-1.5-pro` → `gemini-1.5-flash`) until one is available.
- Validates totals against the invoice grand total — flags rows where the math doesn't match.
- Logs uncertain rows to a separate sheet for manual review (`uncertain_log`) and extraction issues (`extraction_log`).
- Outputs a clean Excel of all trips.

### Stage 2 — `ניסיון לשיפור מסלולים.py` (Route enrichment via Google Maps)
- Reads the extracted trips Excel.
- For each trip: builds a route key (origin → waypoints → destination) and queries **Google Maps Directions API** for distance & duration.
- **Caches** every queried route in `routes_cache.json` to avoid re-paying for the same lookups across runs.
- Fuzzy-matches misspelled / ambiguous locations against a set of known good locations (`difflib.get_close_matches`) and logs suggestions for the human reviewer.
- Produces an enriched Excel with km, minutes per trip, and a separate sheet for problematic locations.

## Tech stack

- **Python 3.11**
- **google-genai** — Gemini API for invoice OCR + extraction
- **requests** — Google Maps Directions API
- **pdfplumber, openpyxl, pandas** — file I/O
- **python-dotenv** — secrets

## Project layout

```
פייתון/
  תיקון לנושא סהכ חשבוניות.py    Stage 1: invoices → structured trips
  ניסיון לשיפור מסלולים.py        Stage 2: trips + Google Maps → enriched dataset
חשבוניות/                       Drop-zone for invoice files to process
חשבוניות ופירוטים אחרי ריצה/    Processed invoices archive
חשבונית לא ברורה/                Invoices flagged for manual review
דוחות/                          Outputs
  נתונים_גולמיים_*.xlsx            Stage 1 raw extraction
  נתונים_מעובדים_*.xlsx            Stage 2 enriched output
  לוג_מיקומים_בעייתיים_*.xlsx     Problematic locations log
  routes_cache.json                Google Maps response cache (excluded from repo)
הסעות מנהיגות - אחרי עיבוד.xlsx  Sample final output
```

## Setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl pdfplumber google-genai requests python-dotenv

# Create .env (copy from .env.example) and fill in your own keys:
#   GEMINI_API_KEY=...
#   GOOGLE_MAPS_API_KEY=...
```

## Usage

```bash
# 1. Drop invoice PDFs/Excels into חשבוניות/

# 2. Extract structured trips from invoices
python "פייתון/תיקון לנושא סהכ חשבוניות.py"

# 3. Enrich with route distance & duration
python "פייתון/ניסיון לשיפור מסלולים.py"
```

## Notes

- All UI, data and reports are in **Hebrew** (RTL).
- API keys are loaded from `.env` (never committed). See `.env.example` for the required variables.
- The `routes_cache.json` is excluded from the repo — it's an on-disk cache that grows over time.
- Hardcoded `BASE_FOLDER` paths inside scripts point to OneDrive folders specific to the user — adjust before running elsewhere.

## Status

Used ad-hoc to process monthly transportation invoice batches.

---

*Author: Dror Nadel · 2026*
