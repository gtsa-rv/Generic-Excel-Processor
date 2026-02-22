# Generic Excel Processor

A Python script that parses multi-sheet Excel files and generates a structured summary report with pricing and availability statistics, grouped by complex and room count.

---

## Features

- Processes multiple named sheets from a single Excel file
- Auto-detects header rows and key columns (price, area, rooms, status, ID)
- Filters apartments by availability status (`вільно`)
- Supports mixed sheets where the complex name is derived from the apartment ID
- Outputs a formatted `.xlsx` summary report with min/max area, price per m², and total price — each annotated with the corresponding apartment ID

---

## Requirements

- Python 3.8+
- `pandas`
- `openpyxl`

Install dependencies:

```bash
pip install pandas openpyxl
```

---

## Configuration

Before running, open `main.py` and fill in the configuration section near the top of the file:

```python
# Sheet names to process from your Excel file
SHEETS_TO_PROCESS = ["Sheet1", "Sheet2"]

# Apartment ID fragments to skip (e.g. placeholder or template rows)
EXCLUDED_ID_MARKERS = ["EXAMPLE", "TEST"]

# For mixed sheets: map ID fragments to complex names
ID_TO_COMPLEX_RULES = [
    {"id_contains": "ABC", "complex_name": "Complex Alpha"},
    {"id_contains": "XYZ", "complex_name": "Complex Beta"},
]

# Map sheet-name fragments to normalized complex names
SHEET_TO_COMPLEX_RULES = [
    {"sheet_contains": "alpha", "complex_name": "Complex Alpha"},
]

# Sheet-name keywords that indicate a mixed sheet (ID-based complex mapping)
MIXED_SHEET_KEYWORDS = ["mixed", "combined"]
```

---

## Usage

```bash
python main.py <input_file> [options]
```

### Arguments

| Argument | Description |
|---|---|
| `input_file` | Path to the input `.xlsx` file |
| `-o`, `--output` | Output file path (default: `Summary_Report.xlsx`) |
| `-v`, `--verbose` | Print detailed per-sheet diagnostics |
| `--preview` | Print the full summary table to the console |

### Examples

```bash
# Basic usage
python main.py data.xlsx

# Custom output file
python main.py data.xlsx -o reports/output.xlsx

# Verbose mode with console preview
python main.py data.xlsx -v --preview
```

---

## Output

The script produces a single-sheet Excel file (`Summary`) with the following columns:

| Column | Description |
|---|---|
| `GROUP` | Complex / building name |
| `ROOMS` | Number of rooms (1, 2, or 3) |
| `MIN_AREA` | Smallest available unit area (m²) |
| `MAX_AREA` | Largest available unit area (m²) |
| `MIN_PRICE_PER_M2` | Lowest price per m² and apartment ID |
| `MAX_PRICE_PER_M2` | Highest price per m² and apartment ID |
| `MIN_TOTAL_PRICE` | Cheapest total price and apartment ID |

Rows are sorted by complex name and room count.

---

## Column Detection Logic

The script uses fuzzy keyword matching to locate the right columns in each sheet. It looks for Ukrainian-language headers by default (e.g. `ціна за метр`, `Площа`, `Статус`, `Розмір`) with fallback patterns for variations. If your source file uses different header names, extend the matching logic in `process_sheet()`.

---

## Notes

- Only apartments with status `вільно` (available) are included.
- Only units with 1, 2, or 3 rooms are processed.
- Rows missing both price per m² and total price are skipped.
- Currency values support `$`, `грн`, commas, and non-breaking spaces.