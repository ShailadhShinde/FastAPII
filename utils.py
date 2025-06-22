import pandas as pd
from io import BytesIO
from typing import Dict, List, Tuple

# The list of exact table titles to anchor on
TABLE_TITLES = [
    "INITIAL INVESTMENT",
    "CASHFLOW DETAILS",
    "DISCOUNT RATE",
    "WORKING CAPITAL",
    "GROWTH RATES",
    "SALVAGE VALUE",
    "OPERATING CASHFLOWS",
    "BOOK VALUE & DEPRECIATION",
    "Investment Measures",
]

def locate_table_titles(df: pd.DataFrame) -> List[Tuple[int, int, str]]:
    """Find all positions of table titles in the spreadsheet"""
    rows, cols = df.shape
    data = df.values
    title_positions = []
    
    for r in range(rows):
        for c in range(cols):
            cell = str(data[r, c]).strip()
            if cell in TABLE_TITLES:
                title_positions.append((r, c, cell))
    
    title_positions.sort()
    return title_positions

def calculate_row_boundaries(title_positions: List[Tuple[int, int, str]], total_rows: int) -> Dict[int, int]:
    """Calculate vertical boundaries for each table"""
    title_rows = []
    for r, c, title in title_positions:
        if r not in title_rows:
            title_rows.append(r)
    title_rows.sort()
    
    row_boundaries = {}
    for i in range(len(title_rows)):
        this_row = title_rows[i]
        if i + 1 < len(title_rows):
            next_row = title_rows[i + 1]
        else:
            next_row = total_rows
        row_boundaries[this_row] = next_row
    
    return row_boundaries

def group_titles_by_row(title_positions: List[Tuple[int, int, str]]) -> Dict[int, List[int]]:
    """Group title columns by their row positions"""
    row_to_cols = {}
    for r, c, title in title_positions:
        if r not in row_to_cols:
            row_to_cols[r] = []
        row_to_cols[r].append(c)
    return row_to_cols

def determine_horizontal_boundaries(start_row: int, start_col: int, row_to_cols: Dict[int, List[int]], total_cols: int) -> Tuple[int, int]:
    """Determine left/right boundaries for a table"""
    cols_on_this_row = row_to_cols[start_row].copy()
    cols_on_this_row.sort()
    
    # Default full width
    left = 0
    right = total_cols

    if len(cols_on_this_row) > 1:
        # Build list of segments between each title column
        segments = []
        for i in range(len(cols_on_this_row)):
            seg_start = cols_on_this_row[i]
            if i + 1 < len(cols_on_this_row):
                seg_end = cols_on_this_row[i + 1]
            else:
                seg_end = total_cols
            segments.append((seg_start, seg_end))
        
        # Find which segment contains our start_col
        for seg_left, seg_right in segments:
            if seg_left <= start_col < seg_right:
                left = seg_left
                right = seg_right
                break
    
    return left, right

def extract_table_block(df: pd.DataFrame, top: int, bottom: int, left: int, right: int) -> pd.DataFrame:
    """Extract and clean a table block from the spreadsheet"""
    block = df.iloc[top:bottom, left:right].copy()
    block = block.replace("", pd.NA)
    
    # Drop empty rows
    block = block.dropna(how="all", axis=0)
    # Drop empty cols
    block = block.dropna(how="all", axis=1)
    
    block = block.reset_index(drop=True)
    return block

def find_tables_in_spreadsheet(df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    """Find and extract all tables from the spreadsheet"""
    rows, cols = df.shape
    title_positions = locate_table_titles(df)
    row_to_cols = group_titles_by_row(title_positions)
    row_boundaries = calculate_row_boundaries(title_positions, rows)
    
    # Track used cells to avoid overlap
    already_used = []
    for r in range(rows):
        row = []
        for c in range(cols):
            row.append(False)
        already_used.append(row)
    
    extracted_tables = {}
    title_counts = {}  # Track counts for duplicate titles
    
    for start_row, start_col, base_title in title_positions:
        if already_used[start_row][start_col]:
            continue
            
        # Handle duplicate titles
        if base_title in title_counts:
            title_counts[base_title] += 1
            title = base_title + " (" + str(title_counts[base_title]) + ")"
        else:
            title_counts[base_title] = 0
            title = base_title
        
        top = start_row
        bottom = row_boundaries[start_row]
        left, right = determine_horizontal_boundaries(start_row, start_col, row_to_cols, cols)
        
        block = extract_table_block(df, top, bottom, left, right)
        
        # Only keep tables with meaningful data
        if block.shape[0] >= 2 and block.shape[1] >= 2:
            extracted_tables[title] = block
            
            # Mark the area as used
            for rr in range(top, bottom):
                for cc in range(left, right):
                    if rr < len(already_used) and cc < len(already_used[0]):
                        already_used[rr][cc] = True
    
    return extracted_tables

def filter_rows_columns(df: pd.DataFrame, max_empty_row_ratio: float, max_empty_col_ratio: float) -> pd.DataFrame:
    """Filter out mostly empty rows and columns"""
    # Row filtering
    na_counts_per_row = df.isna().mean(axis=1)
    keep_rows = []
    for idx in range(len(na_counts_per_row)):
        if na_counts_per_row.iloc[idx] <= max_empty_row_ratio:
            keep_rows.append(idx)
    
    df = df.iloc[keep_rows, :]
    
    # Column filtering
    na_counts_per_col = df.isna().mean(axis=0)
    keep_cols = []
    for col in df.columns:
        if na_counts_per_col[col] <= max_empty_col_ratio:
            keep_cols.append(col)
    
    df = df[keep_cols]
    return df

def trim_blank_edges(data: pd.DataFrame, by_rows: bool, max_blank_streak: int) -> pd.DataFrame:
    """Trim blank runs at edges of the table"""
    if by_rows:
        pattern = data.isna().all(axis=1).tolist()
        labels = list(data.index)
    else:
        pattern = data.isna().all(axis=0).tolist()
        labels = list(data.columns)
    
    # Find trim points
    start = 0
    end = len(pattern) - 1
    
    # Trim leading blanks
    run = 0
    for i in range(len(pattern)):
        is_blank = pattern[i]
        if is_blank:
            run += 1
        else:
            break
        if run > max_blank_streak:
            start = i + 1
    
    # Trim trailing blanks
    run = 0
    reversed_pattern = list(reversed(pattern))
    for i in range(len(reversed_pattern)):
        is_blank = reversed_pattern[i]
        if is_blank:
            run += 1
        else:
            break
        if run > max_blank_streak:
            end = len(pattern) - 1 - (i + 1)
    
    selected = labels[start : end + 1]
    if by_rows:
        return data.loc[selected, :]
    else:
        return data.loc[:, selected]

def clean_up_table(
    table: pd.DataFrame,
    max_empty_row_ratio: float = 0.6,
    max_empty_col_ratio: float = 0.6,
    max_blank_streak: int = 1
) -> pd.DataFrame:
    """Clean up table by removing empty rows/columns and trimming edges"""
    df = table.copy()
    df = filter_rows_columns(df, max_empty_row_ratio, max_empty_col_ratio)
    df = trim_blank_edges(df, by_rows=True, max_blank_streak=max_blank_streak)
    df = trim_blank_edges(df, by_rows=False, max_blank_streak=max_blank_streak)
    df = df.reset_index(drop=True)
    return df

def convert_table_to_list(table: pd.DataFrame) -> List[List[str]]:
    """Convert DataFrame to list of lists of strings"""
    result = []
    for row in table.values.tolist():
        string_row = []
        for cell in row:
            string_row.append(str(cell))
        result.append(string_row)
    return result

def process_excel_file(excel_bytes: bytes) -> Dict[str, List[List[str]]]:
    """Process Excel file and return cleaned tables as lists"""
    # Read into DataFrame
    bio = BytesIO(excel_bytes)
    df = pd.read_excel(bio, header=None, dtype=str, engine="xlrd").fillna("")
    
    # Extract and clean tables
    raw_tables = find_tables_in_spreadsheet(df)
    final_tables = {}
    
    for title, block in raw_tables.items():
        cleaned = clean_up_table(block)
        final_tables[title] = convert_table_to_list(cleaned)
    
    return final_tables