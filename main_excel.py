import xlwings as xw
import math

# Constants for A4 dimensions in cm
A4_SIZES = {
    'portrait':  {'width_cm': 21.0, 'height_cm': 29.7},
    'landscape': {'width_cm': 29.7, 'height_cm': 21.0},
}

# Conversion factors
POINTS_TO_CM = 0.03528       # 1 point = 0.03528 cm
EXCEL_UNIT_TO_CM = 0.142     # approx for Calibri 11 at 100% zoom


def split_range_into_pages(
    file_path: str,
    sheet_name: str,
    range_address: str,
    orientation: str = 'landscape'
):
    """
    Splits the given Excel range into page‐sized sub‐ranges for A4 printing.
    
    Returns a list of (start_row, end_row) tuples **within** the original range,
    and their corresponding A1 addresses.
    """
    # Open workbook and sheet
    wb = xw.Book(file_path)
    sht = wb.sheets[sheet_name]
    rng = sht.range(range_address)
    
    # 1) Gather original column widths & row heights
    col_widths_csv = [c.column_width * EXCEL_UNIT_TO_CM for c in rng.columns]
    row_heights_cm   = [r.row_height * POINTS_TO_CM     for r in rng.rows]
    
    total_width_cm = sum(col_widths_csv)
    page_w_cm = A4_SIZES[orientation]['width_cm']
    page_h_cm = A4_SIZES[orientation]['height_cm']
    
    # 2) Compute scale so width fits exactly
    scale = page_w_cm / total_width_cm
    
    # 3) Determine the max original‐range height (in cm)
    #    that when scaled fits into the page height
    max_original_height_cm = page_h_cm / scale
    
    # 4) Partition rows
    pages = []
    start_idx = 0
    n_rows = len(row_heights_cm)
    
    while start_idx < n_rows:
        acc = 0.0
        end_idx = start_idx
        # add rows until adding the next would overflow
        while end_idx < n_rows and (acc + row_heights_cm[end_idx]) <= max_original_height_cm:
            acc += row_heights_cm[end_idx]
            end_idx += 1
        
        # if the very first row is bigger than a page, force at least one row
        if end_idx == start_idx:
            end_idx += 1
        
        pages.append((start_idx, end_idx - 1))
        start_idx = end_idx
    
    # 5) Convert partitions back to A1 addresses
    page_ranges = []
    top_row = rng.row
    left_col = rng.column
    for (r0, r1) in pages:
        start_cell = sht.cells[top_row - 1 + r0, left_col - 1]
        end_cell   = sht.cells[top_row - 1 + r1, left_col - 1 + rng.columns.count - 1]
        page_ranges.append(f"{start_cell.address}:{end_cell.address}")
    
    wb.close()
    return page_ranges


# Example usage:
if __name__ == "__main__":
    pages = split_range_into_pages(
        file_path="./Itut 7A Well Completions Schematic.xls",
        sheet_name="Completion String",
        range_address="B2:AD88",
        orientation="landscape"
    )
    for i, pr in enumerate(pages, start=1):
        print(f"Page {i}: {pr}")