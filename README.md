# LedgerFlow

Processes bank transaction Excel exports, categorizes them, and writes the
results into a yearly Excel workbook.

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

## Setup

```bash
pip install -e ".[dev]"       # core + test dependencies
pip install -e ".[dev,gpt]"   # include openai for GPT categorization
pytest
```

---

## Configuration (`config.json`)

```json
{
  "include_mortgage": false,
  "seed_mapping_path": "data/seed_categories.xlsx",
  "cache_path": "data/category_cache.json",
  "dedup_amount_rounding": 0
}
```

All keys are optional; defaults are used for anything omitted.

---

## Categories

`Other` · `Business` · `Pharm` · `Insurance` · `Transportation` · `Leisure` · `Home`

`Mortgage` is available when `include_mortgage: true` in config.
