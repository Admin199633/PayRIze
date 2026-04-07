# LedgerFlow


התקנה והגדרה
1️⃣ הורדת הפרויקט
הורד את הקובץ ZIP מ־GitHub
פתח את ה‑ZIP ו־העתק את התיקייה למיקום הרצוי, לדוגמה:
C:\Users\LiorSw\Downloads
פתח CMD במיקום הזה:
cd C:\Users\LiorSw\Downloads
שנה את שם התיקייה (כדי להוריד את -main):
ren LedgerFlow-main LedgerFlow
2️⃣ התקנת Python (אם לא מותקן)
winget install Python.Python.3.12

לאחר ההתקנה, בדוק שהתקנה עברה בהצלחה:

python --version

או

py --version
3️⃣ התקנת חבילות נדרשות

ב־CMD, בתוך התיקייה הראשית של הפרויקט:

cd LedgerFlow
py -m pip install -r requirements.txt

אם אין requirements.txt, התקן חבילות נדרשות לפי הקוד שלך (למשל PyInstaller, google-api-python-client, google-auth).

4️⃣ הגדרת Google Sheets (credentials)
פתח את Google Cloud Console
צור Service Account חדש
צור JSON Key עבור ה‑Service Account
הורד את הקובץ ושנה את שמו ל:
credentials.json
העתק את הקובץ לתיקייה הראשית של הפרויקט:
C:\Users\LiorSw\Downloads\LedgerFlow\credentials.json
שתף את Google Sheet עם ה‑email של ה‑Service Account (Editor).
5️⃣ הרצת הפרויקט

בתוך התיקייה הראשית של הפרויקט:

py -m ledgerflow

או אם אתה בתוך ledgerflow:

py __main__.py
6️⃣ לבנות EXE (אופציונלי)

אם רוצים ליצור קובץ הפעלה:

py -m PyInstaller ledgerflow/__main__.py

או אם יש .spec מוכן:

py -m PyInstaller LedgerFlow.spec


Processes bank transaction Excel exports, categorizes them, and writes the
results into a yearly Excel workbook.

---

## First-run setup (Windows)

`pyproject.toml` is the dependency source of truth for this repository.

1. Create a virtual environment:

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

2. Install the app and optional extras you need:

```powershell
python -m pip install -e .
python -m pip install -e ".[gpt]"     # only if you want GPT categorization
python -m pip install -e ".[build]"   # only if you want to build the Windows .exe
```

3. Create your local config:

```powershell
Copy-Item config.json.example config.json
```

4. Put your Google service-account file here:

```text
LedgerFlow\credentials.json
```

5. Edit `config.json` and set:
- `seed.spreadsheet_id`
- `seed.worksheet_name`
- `year.spreadsheet_id`
- `year.month_sheets`
- `google_service_account_file`
- `currency_rates`

6. Run the desktop app:

```powershell
python -m ledgerflow
```

### Fresh machine checklist

- Create and activate a venv
- Install dependencies from `pyproject.toml`
- Copy `config.json.example` to `config.json`
- Place `credentials.json` next to `config.json`
- Fill in the Google Sheets IDs and month-sheet mapping
- Set manual currency rates in `config.json` or in the app Config tab
- Launch `python -m ledgerflow`

---

## Architecture

```
LedgerFlow/
├── ledgerflow/
│   ├── models.py           # Domain entities: Transaction, Category, DedupKey
│   ├── config.py           # Config dataclass; loaded from JSON or defaults
│   ├── dedup.py            # Dedup key generation and duplicate removal
│   ├── extraction.py       # Excel → Transaction list
│   ├── categorization.py   # Seed/cache categorization; CategorizationResult
│   ├── gpt_categorizer.py  # GPT batch categorization (openai optional dep)
│   ├── filters.py          # filter_by_month (pure)
│   ├── csv_writer.py       # write_raw_csv / write_categorized_csv
│   ├── pipeline.py         # run_pipeline / run_pipeline_from_config
│   └── utils/
│       ├── normalization.py  # Merchant name canonicalization
│       ├── dates.py          # Multi-format date parsing
│       └── amounts.py        # Amount parsing and dedup rounding
└── tests/                    # 330 tests, all passing
```

### Key design decisions

| Decision | Rationale |
|---|---|
| Plain `dataclass` for `Transaction` | No framework coupling; easy to construct in tests. |
| `normalized_merchant` stored on the model | Normalization runs once at extraction time; all downstream code uses the stable form. |
| `DedupKey = tuple[date, str, Decimal]` | Plain tuple → hashable; no extra class needed. |
| `Config` loaded from JSON | Locale-specific header aliases and stop tokens stay out of code. |
| `Category` as `str` Enum | Values serialize to readable strings in CSV / JSON without extra steps. |
| `Mortgage` opt-in via config | The category only appears when the yearly workbook structure supports it. |
| Utilities in `utils/` subpackage | Pure functions with no cross-module imports; trivially testable and reusable. |
| `_extract_from_sheet` is pure (no openpyxl types) | Sheet tests run with plain Python lists — no temp files, no I/O, very fast. |
| `detect_table_header` takes a single row | Fits the state-machine; column positions update per-row, enabling repeated tables. |
| Stop-row detection is substring-based over any cell | Simple and catches all common summary/total patterns; configurable via `stop_tokens`. |

### Data flow

```
Excel files
    └─► extraction.extract_from_files()
            │  (detect headers, stop at summary rows, parse rows)
            ▼
    list[Transaction]  (raw, uncategorized)
            │
            ▼
    dedup.deduplicate()
            │
            ▼
    filter by target month
            │
            ▼
    categorization.categorize_transactions()
            │  seed mapping → cache → GPT batch → Other fallback
            ▼
    list[Transaction]  (categorized)
            │
            ├─► write raw CSV
            ├─► write categorized CSV
            └─► write yearly Excel workbook (month sheet + category block)
```

---

## Extraction notes

- Supports `.xlsx` and `.xlsm`.  `.xls` files raise `UnsupportedFileTypeError` with a clear message to re-save in Excel.
- Each sheet is scanned independently.  Rows before the first valid header are silently ignored.
- Multiple tables per sheet are supported: a new header row restarts the column positions.
- A table ends on: a stop-token row, an all-empty row, or a new header row.
- Stop tokens are matched as case-insensitive substrings.  Known limitation: a merchant name that *contains* a stop token (e.g. "Total Sports") would prematurely end the table.  Work-around: remove the conflicting token from `config.stop_tokens`.

## Development setup

```powershell
python -m pip install -e ".[dev]"
python -m pip install -e ".[dev,gpt]"
pytest
```

## Windows EXE build

Use PyInstaller in windowed mode so LedgerFlow launches as a desktop app without
a console window.

### Build command

```powershell
pip install -e ".[build]"
python -m PyInstaller --noconfirm LedgerFlow.spec
```

Or run:

```powershell
.\build_exe.ps1
```

### Output

The executable is created here:

```text
dist\LedgerFlow\LedgerFlow.exe
```

Double-click [LedgerFlow.exe](/C:/Users/liorsw/Downloads/LedgerFlow/dist/LedgerFlow/LedgerFlow.exe)
after building.

### External files next to the EXE

Keep these files external and editable in the same folder as the executable:

```text
dist\LedgerFlow\config.json
dist\LedgerFlow\credentials.json
```

`config.json` is loaded from next to the `.exe`.

If `config.json` contains:

```json
{
  "google_service_account_file": "credentials.json"
}
```

then `credentials.json` will also be resolved relative to that same folder.

---

## Path resolution

All file paths in `config.json` (`google_service_account_file`, `cache_path`,
`seed_mapping_path`) are resolved **relative to the directory that contains
`config.json`** — not the working directory or the Python/EXE launch path.

- **Relative path** → resolved against the folder of `config.json`
- **Absolute path** → used as-is

This means you can move the entire project folder (or copy it to a new machine)
without changing any paths.  The only requirement is that `credentials.json`,
`cache.json`, etc. sit next to `config.json` — or that you use absolute paths if
they live elsewhere.

The resolved absolute path is written to the log at startup so you can always
verify which file is actually being used.

---

## Configuration (`config.json`)

```json
{
  "seed": {
    "spreadsheet_id": "YOUR_SEED_SPREADSHEET_ID",
    "worksheet_name": "Seed"
  },
  "year": {
    "spreadsheet_id": "YOUR_YEAR_SPREADSHEET_ID",
    "month_sheets": {
      "3": "March"
    }
  },
  "run_log_worksheet_name": "RunLog",
  "google_service_account_file": "credentials.json",
  "currency_rates": {
    "USD": 3.149,
    "EUR": 3.92,
    "GBP": 4.61
  },
  "cache_path": "cache.json"
}
```

Start from `config.json.example`. Use placeholders only in the example; do not
commit real spreadsheet IDs or credentials.

---

## Categories

`Other` · `Business` · `Pharm` · `Insurance` · `Transportation` · `Leisure` · `Home`

`Mortgage` is available when `include_mortgage: true` in config.
"# PayRIze" 
