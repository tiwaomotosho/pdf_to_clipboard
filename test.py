import xlwings as xw

def get_range_dimensions(file_path, sheet_name, range_address):
    
    wb = xw.Book(file_path)
    sht = wb.sheets[sheet_name]
    rng = sht.range(range_address)

    total_width_units = sum(col.column_width for col in rng.columns)
    total_height_points = sum(row.row_height for row in rng.rows)

    # Convert
    total_width_cm = total_width_units * 0.142
    total_height_cm = total_height_points * 0.03528

    print(f"Range {range_address} in {sheet_name}:")
    print(f"Total Width ≈ {total_width_cm:.2f} cm")
    print(f"Total Height ≈ {total_height_cm:.2f} cm")

    wb.close()

# Example
get_range_dimensions("C:/Users/t9omoto\Repos/pdf_to_clipboard/Itut 7A Well Completions Schematic.xls", "Completion String", "B2:AD88")