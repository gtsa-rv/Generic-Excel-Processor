"""
Generic Excel Processor
Parses multi-sheet Excel files and builds a summary table.
"""

import pandas as pd
import re
import warnings
import argparse
import os
from typing import List, Dict, Optional
warnings.filterwarnings('ignore')


# ============================================================================
# PROJECT CONFIGURATION (replace placeholders before running)
# ============================================================================

# Replace these with your actual sheet names.
SHEETS_TO_PROCESS = []

# If an apartment ID contains any of these markers, that row will be skipped.
# Replace/remove based on your own business rules.
EXCLUDED_ID_MARKERS = []

# For mixed sheets, map ID markers to normalized complex names.
# Fill with your own ID fragments and target names.
ID_TO_COMPLEX_RULES = []

# Map sheet-name fragments to normalized complex names.
# Fill with your own sheet-name keywords and target names.
SHEET_TO_COMPLEX_RULES = []

# Sheet-name keywords that indicate a mixed sheet requiring ID-based mapping.
MIXED_SHEET_KEYWORDS = []


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def normalize_text(text) -> str:
    """Normalize text for better column matching"""
    if pd.isna(text):
        return ""
    text = str(text).lower().strip()
    # Remove extra spaces
    text = re.sub(r'\s+', ' ', text)
    return text


def find_best_column(df: pd.DataFrame, candidates_list: List[List[str]],
                     exclude_keywords: List[str] = None) -> Optional[str]:
    """
    Find the best matching column name based on priority keywords.

    Args:
        df: DataFrame to search
        candidates_list: List of keyword lists, ordered by priority
        exclude_keywords: Keywords to exclude from matches

    Returns:
        Column name if found, None otherwise
    """
    exclude_keywords = exclude_keywords or []

    # Normalize column names
    col_map = {normalize_text(col): col for col in df.columns}

    # Try each priority level
    for candidates in candidates_list:
        for col_norm, col_orig in col_map.items():
            # Check if any candidate keyword is in column name
            if any(normalize_text(keyword) in col_norm for keyword in candidates):
                # Check exclusions
                if exclude_keywords and any(normalize_text(excl) in col_norm for excl in exclude_keywords):
                    continue
                return col_orig

    return None


def clean_currency(val) -> Optional[float]:
    """Clean currency values and convert to float"""
    if pd.isna(val):
        return None

    s = str(val).replace(' ', '').replace('$', '').replace('грн', '')
    s = s.replace(',', '.').replace('\xa0', '')

    try:
        return float(s)
    except:
        return None


def extract_room_count(val) -> Optional[int]:
    """Extract room count from string (e.g., '1M' -> 1, '2к' -> 2)"""
    if pd.isna(val):
        return None

    # Look for first digit
    match = re.search(r'(\d+)', str(val))
    if match:
        rooms = int(match.group(1))
        # Only return 1, 2, or 3
        if rooms in [1, 2, 3]:
            return rooms

    return None


def find_header_row(df: pd.DataFrame, keywords: List[str] = None) -> int:
    """
    Find the header row index by looking for key column names.

    Args:
        df: DataFrame with potentially merged cells
        keywords: Keywords to search for (default: ['ID', 'Статус'])

    Returns:
        Index of header row
    """
    keywords = [normalize_text(kw) for kw in (keywords or ['ID', 'Статус', 'Площа', 'ціна'])]

    for i, row in df.iterrows():
        normalized_cells = [normalize_text(cell) for cell in row.tolist()]
        if any(any(kw in cell for kw in keywords) for cell in normalized_cells):
            return i

    return 0  # Default to first row


# ============================================================================
# MAIN PROCESSING FUNCTIONS
# ============================================================================

def process_sheet(df: pd.DataFrame, sheet_name: str, verbose: bool = False) -> List[Dict]:
    """
    Process a single sheet and extract standardized data.

    Args:
        df: DataFrame from Excel sheet
        sheet_name: Name of the sheet

    Returns:
        List of dictionaries with standardized data
    """

    # Find header row
    header_idx = find_header_row(df)

    # Set headers
    df.columns = df.iloc[header_idx]
    df = df[header_idx + 1:].copy()

    # Reset index
    df = df.reset_index(drop=True)

    if verbose:
        print(f"  Processing sheet: {sheet_name}")
        print(f"  Columns found: {list(df.columns)[:10]}...")

    # ========================================================================
    # COLUMN IDENTIFICATION (PRIORITY ORDER)
    # ========================================================================

    # 1. Price per meter - Look for "ціна за метр" (standard price, NOT discount)
    price_m2_col = None
    for col in df.columns:
        col_str = str(col).strip()
        col_norm = normalize_text(col)
        # Priority 1: Exact "ціна за метр" (with possible trailing space)
        if col_norm == 'ціна за метр' or col_str == 'ціна за метр ':
            price_m2_col = col
            break

    # Fallback: For different sheets
    if not price_m2_col:
        for col in df.columns:
            col_norm = normalize_text(col)
            # Example fallback pattern with extra notes in the source header text.
            if 'ціна за 1 м' in col_norm and 'видаляти' in col_norm:
                price_m2_col = col
                break
            # Generic sale-price-per-meter fallback.
            elif ('ціна' in col_norm and 'м' in col_norm and 'продаж' in col_norm):
                if 'базов' not in col_norm and 'старт' not in col_norm:
                    price_m2_col = col
                    break

    # 2. Total Price - Look for "Вартість для продажу" (standard, NOT discount)
    total_price_col = None
    for col in df.columns:
        col_norm = normalize_text(col)
        # Priority: "Вартість для продажу"
        if 'вартість для продажу' in col_norm:
            total_price_col = col
            break

    # Fallback pattern with extra notes in the source header text.
    if not total_price_col:
        for col in df.columns:
            col_norm = normalize_text(col)
            if 'вартість' in col_norm and 'видаляти' in col_norm:
                total_price_col = col
                break
            # Last fallback: "Вартість ПРОДАЖУ"
            elif 'вартість продажу' in col_norm:
                total_price_col = col
                break

    # 4. Rooms - Must be exact "Розмір" not "Націнка за розмір"
    rooms_col = None
    for col in df.columns:
        col_str = str(col).strip()
        # Exact match or close variants
        if col_str in ['Розмір', 'Rooms', 'К-сть кімнат', 'Кімнат']:
            rooms_col = col
            break
        # If not found, try lowercase match
        if normalize_text(col) == 'розмір':
            rooms_col = col
            break

    # 5. Area - Must be exact "Площа" not "площа оновлена"
    area_col = None
    for col in df.columns:
        col_norm = normalize_text(col)
        if col_norm == 'площа':  # Exact match only
            area_col = col
            break

    # 6. Status - Look for "Статус" column specifically
    status_col = find_best_column(df, [
        ['статус'],
        ['стан', 'state'],
    ])

    # 7. ID
    id_col = find_best_column(df, [
        ['id', 'номер'],
    ])

    # 8. Complex (ЖК)
    complex_col = find_best_column(df, [
        ['жк', 'complex', 'building'],
    ])

    if verbose:
        print(f"  Found columns:")
        print(f"    Price/m2 (100%): {price_m2_col}")
        print(f"    Total Price: {total_price_col}")
        print(f"    Rooms: {rooms_col}")
        print(f"    Area: {area_col}")
        print(f"    Status: {status_col}")
        print(f"    Complex: {complex_col}")

    # ========================================================================
    # DATA EXTRACTION
    # ========================================================================

    results = []

    for idx, row in df.iterrows():
        # Get status
        status_val = str(row.get(status_col, '')).lower() if status_col else ''

        # Filter: Only 'вільно' (available)
        if 'вільно' not in status_val:
            continue

        # Get room count
        rooms = extract_room_count(row.get(rooms_col)) if rooms_col else None
        if rooms not in [1, 2, 3]:
            continue

        # Get area
        area = pd.to_numeric(row.get(area_col), errors='coerce') if area_col else None
        if pd.isna(area) or area <= 0:
            continue

        # Get prices
        price_m2 = clean_currency(row.get(price_m2_col)) if price_m2_col else None
        total_price = clean_currency(row.get(total_price_col)) if total_price_col else None

        # Skip if both prices are missing
        if price_m2 is None and total_price is None:
            continue

        # Get complex name
        if complex_col and not pd.isna(row.get(complex_col)):
            complex_name = str(row.get(complex_col)).strip()
        else:
            # Use sheet name as complex name
            complex_name = sheet_name

        # Get ID
        apt_id = str(row.get(id_col, '')).strip() if id_col else ''

        # Skip rows by configured ID markers.
        if any(marker.upper() in apt_id.upper() for marker in EXCLUDED_ID_MARKERS):
            continue

        # Determine complex from ID if this is a mixed sheet.
        if any(keyword.lower() in sheet_name.lower() for keyword in MIXED_SHEET_KEYWORDS):
            apt_id_upper = apt_id.upper()
            for rule in ID_TO_COMPLEX_RULES:
                if rule['id_contains'].upper() in apt_id_upper:
                    complex_name = rule['complex_name']
                    break

        # Standardize complex names from sheet-name rules.
        for rule in SHEET_TO_COMPLEX_RULES:
            if rule['sheet_contains'].lower() in sheet_name.lower():
                complex_name = rule['complex_name']
                break

        results.append({
            'GROUP': complex_name,
            'ROOMS': rooms,
            'ID': apt_id,
            'AREA': area,
            'PRICE_PER_M2': price_m2,
            'TOTAL_PRICE': total_price,
        })

    if verbose:
        print(f"  Extracted {len(results)} available apartments")

    return results


def generate_summary(data: List[Dict]) -> pd.DataFrame:
    """
    Generate summary pivot table from extracted data.

    Args:
        data: List of apartment dictionaries

    Returns:
        Summary DataFrame
    """
    if not data:
        return pd.DataFrame()

    df = pd.DataFrame(data)

    # Remove rows with missing critical values
    df = df.dropna(subset=['PRICE_PER_M2', 'TOTAL_PRICE'])

    results = []

    # Group by Complex and Room Count
    for (complex_name, room_count), group in df.groupby(['GROUP', 'ROOMS']):

        # Find extremes with their IDs
        min_area_row = group.loc[group['AREA'].idxmin()]
        max_area_row = group.loc[group['AREA'].idxmax()]

        min_price_row = group.loc[group['PRICE_PER_M2'].idxmin()]
        max_price_row = group.loc[group['PRICE_PER_M2'].idxmax()]

        cheapest_row = group.loc[group['TOTAL_PRICE'].idxmin()]

        results.append({
            'GROUP': complex_name,
            'ROOMS': str(room_count),
            'MIN_AREA': f"{min_area_row['AREA']:.1f}",
            'MAX_AREA': f"{max_area_row['AREA']:.1f}",
            'MIN_PRICE_PER_M2': f"{min_price_row['PRICE_PER_M2']:,.0f} ({min_price_row['ID']})",
            'MAX_PRICE_PER_M2': f"{max_price_row['PRICE_PER_M2']:,.0f} ({max_price_row['ID']})",
            'MIN_TOTAL_PRICE': f"{cheapest_row['TOTAL_PRICE']:,.0f} ({cheapest_row['ID']})"
        })

    summary = pd.DataFrame(results)

    # Sort by Complex and Room count
    summary = summary.sort_values(by=['GROUP', 'ROOMS'])

    return summary


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main(
    file_path: str,
    output_path: str = 'Summary_Report.xlsx',
    verbose: bool = False,
    preview: bool = False
):
    """
    Main execution function.

    Args:
        file_path: Path to input Excel file
        output_path: Path for output Excel file
    """

    print("GENERIC EXCEL PROCESSOR")

    all_data = []

    # Sheets to process (configured at file top)
    sheets_to_process = SHEETS_TO_PROCESS
    if not sheets_to_process:
        print("❌ SHEETS_TO_PROCESS is empty. Fill configuration values in main.py first.")
        return

    if not os.path.exists(file_path):
        print(f"❌ Input file not found: {file_path}")
        return

    # Load Excel file
    try:
        xl_file = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"❌ Error loading Excel file: {e}")
        return

    for sheet_name in sheets_to_process:
        if sheet_name not in xl_file.sheet_names:
            print(f"⚠️  Sheet '{sheet_name}' not found, skipping...")
            continue

        if verbose:
            print(f"\n{'='*80}")
            print(f"Processing: {sheet_name}")
            print(f"{'='*80}")
        else:
            print(f"Processing: {sheet_name}...")

        try:
            # Read sheet
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

            # Process sheet
            sheet_data = process_sheet(df, sheet_name, verbose=verbose)

            # Add to all data
            all_data.extend(sheet_data)
            print(f"  -> {len(sheet_data)} available apartments")

        except Exception as e:
            print(f"❌ Error processing sheet '{sheet_name}': {e}")
            if verbose:
                import traceback
                traceback.print_exc()

    # ========================================================================
    # GENERATE SUMMARY REPORT
    # ========================================================================

    print("\nGENERATING SUMMARY REPORT")

    if not all_data:
        print("❌ No data extracted. Cannot generate report.")
        return

    print(f"Total apartments found: {len(all_data)}")

    # Generate summary
    summary = generate_summary(all_data)

    if summary.empty:
        print("❌ Summary is empty. Cannot generate report.")
        return

    print(f"Summary rows: {len(summary)}")
    if preview:
        print("\nPreview:")
        print(summary.to_string(index=False))

    # ========================================================================
    # SAVE TO EXCEL
    # ========================================================================

    print(f"Saving to: {output_path}")

    try:
        # Save with formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            summary.to_excel(writer, sheet_name='Summary', index=False)

            # Auto-adjust column widths
            worksheet = writer.sheets['Summary']
            for idx, col in enumerate(summary.columns):
                max_length = max(
                    summary[col].astype(str).apply(len).max(),
                    len(col)
                ) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)

        print("✅ Report saved successfully!")

    except Exception as e:
        print(f"❌ Error saving report: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    print("="*80)
    print("GENERIC EXCEL PROCESSOR")
    print("="*80)

    parser = argparse.ArgumentParser(
        description="Analyze Excel data and build a summary report."
    )
    parser.add_argument(
        "input_file",
        help="Path to the input Excel file."
    )
    parser.add_argument(
        "-o",
        "--output",
        default="Summary_Report.xlsx",
        help="Path to the output Excel file (default: Summary_Report.xlsx)."
    )
    parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Print detailed per-sheet diagnostics."
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Print full summary table to console."
    )

    args = parser.parse_args()
    main(args.input_file, args.output, verbose=args.verbose, preview=args.preview)