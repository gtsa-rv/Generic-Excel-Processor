# Generic Excel Processor

A Python script that parses multi-sheet Excel files and generates a structured summary report with pricing and availability statistics, grouped by complex and room count.

---

## Features

- Processes multiple named sheets from a single Excel file
- Auto-detects header rows and key columns (price, area, rooms, status, ID)
- Filters apartments by availability status (`available`)
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

## Expected Input Format

Your Excel file should contain sheets with headers that include some of the following column names (exact names or close variants):

| Data | Expected Header Examples |
|---|---|
| Price per m² | `price per meter`, `price per 1 m (sale)` |
| Total price | `sale price`, `total price` |
| Room count | `Rooms`, `Room Count`, `Room`, `Size` |
| Area | `Area` |
| Status | `Status`, `State`, `Availability` |
| Unit ID | `ID`, `Number`, `Unit` |
| Complex | `Complex`, `Building`, `Project` |

The script will auto-detect the header row even if the first few rows contain merged cells or metadata.

---

## Output

The script produces a single-sheet Excel file (`Summary`) with the following columns:

| Column | Description |
|---|---|
| `GROUP` | Complex / building name |
| `ROOMS` | Number of rooms (1, 2, or 3) |
| `MIN_AREA` | Smallest available unit area (m²) |
| `MAX_AREA` | Largest available unit area (m²) |
| `MIN_PRICE_PER_M2` | Lowest price per m² and unit ID |
| `MAX_PRICE_PER_M2` | Highest price per m² and unit ID |
| `MIN_TOTAL_PRICE` | Cheapest total price and unit ID |

Rows are sorted by complex name and room count.

---

## Notes

- Only units with status `available` are included.
- Only units with 1, 2, or 3 rooms are processed.
- Rows missing both price per m² and total price are skipped.
- Currency values support `$`, `uah`, commas, and non-breaking spaces.
- If no `Complex` / `Building` column is found in a sheet, the sheet name is used as the group name.