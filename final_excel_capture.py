import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import xlwings as xw
import fitz  # PyMuPDF
from PIL import Image
import win32clipboard
import io
import os
import sys
import tempfile

# Constants for A4 dimensions in cm
A4_SIZES = {
    'portrait':  {'width_cm': 21.0, 'height_cm': 29.7},
    'landscape': {'width_cm': 29.7, 'height_cm': 21.0},
}

# Conversion factors
POINTS_TO_CM = 0.03528       # 1 point = 0.03528 cm
EXCEL_UNIT_TO_CM = 0.142     # Approx for Calibri 11 at 100% zoom

# Quality to DPI mapping
QUALITY_TO_DPI = {
    "Low Quality": 100,
    "Medium Quality": 300,
    "High Quality": 600
}

# Global variable to store page ranges
pages = []

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_sheet_names(file_path):
    """Retrieve sheet names from the selected Excel file."""
    app = xw.App(visible=False)
    wb = app.books.open(file_path)
    sheet_names = [sheet.name for sheet in wb.sheets]
    wb.close()
    app.quit()
    return sheet_names

def split_range_into_pages(temp_file_path: str, sheet_name: str, range_address: str, orientation: str = 'landscape'):
    """
    Splits the given Excel range into page-sized sub-ranges for A4 printing using a temporary file.
    Returns a list of A1 address ranges for each page.
    """
    wb = xw.Book(temp_file_path)
    sht = wb.sheets[sheet_name]
    rng = sht.range(range_address)
    
    col_widths_cm = [c.column_width * EXCEL_UNIT_TO_CM for c in rng.columns]
    row_heights_cm = [r.row_height * POINTS_TO_CM for r in rng.rows]
    
    total_width_cm = sum(col_widths_cm)
    page_w_cm = A4_SIZES[orientation]['width_cm']
    page_h_cm = A4_SIZES[orientation]['height_cm']
    
    scale = page_w_cm / total_width_cm
    max_original_height_cm = page_h_cm / scale
    
    pages_list = []
    start_idx = 0
    n_rows = len(row_heights_cm)
    
    while start_idx < n_rows:
        acc = 0.0
        end_idx = start_idx
        while end_idx < n_rows and (acc + row_heights_cm[end_idx]) <= max_original_height_cm:
            acc += row_heights_cm[end_idx]
            end_idx += 1
        if end_idx == start_idx:
            end_idx += 1
        pages_list.append((start_idx, end_idx - 1))
        start_idx = end_idx
    
    page_ranges = []
    top_row = rng.row
    left_col = rng.column
    for (r0, r1) in pages_list:
        start_cell = sht.cells[top_row - 1 + r0, left_col - 1]
        end_cell = sht.cells[top_row - 1 + r1, left_col - 1 + rng.columns.count - 1]
        page_ranges.append(f"{start_cell.address}:{end_cell.address}")
    
    wb.close()
    return page_ranges

def export_range_to_pdf(temp_file_path, sheet_name, range_address, orientation):
    """Export the specified Excel range to a temporary PDF with correct orientation and scaling."""
    app = xw.App(visible=False)
    wb = app.books.open(temp_file_path)
    sht = wb.sheets[sheet_name]
    
    sht.api.PageSetup.PrintArea = range_address
    
    if orientation == "landscape":
        sht.api.PageSetup.Orientation = 2  # xlLandscape
    else:
        sht.api.PageSetup.Orientation = 1  # xlPortrait
    
    sht.api.PageSetup.Zoom = False
    sht.api.PageSetup.FitToPagesWide = 1
    sht.api.PageSetup.FitToPagesTall = 1
    
    temp_pdf = os.path.join(tempfile.gettempdir(), "temp_excel_page.pdf")
    sht.to_pdf(temp_pdf)
    
    wb.close()
    app.quit()
    return temp_pdf

def capture_page(temp_file_path, sheet_name, range_address, orientation, dpi, crop_ratio):
    """Capture the specified range as an image with given DPI, crop it, and copy to clipboard."""
    temp_pdf = export_range_to_pdf(temp_file_path, sheet_name, range_address, orientation)
    
    doc = fitz.open(temp_pdf)
    page = doc.load_page(0)
    zoom = dpi / 72
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=False)
    img_data = pix.tobytes("png")
    
    image = Image.open(io.BytesIO(img_data))
    
    # Crop the image if crop_ratio is less than 1
    if crop_ratio < 1.0:
        width, height = image.size
        new_height = int(height * crop_ratio)
        if new_height < 1:  # Prevent cropping to zero height
            new_height = 1
        image = image.crop((0, 0, width, new_height))  # Retain top part
    
    # Copy to clipboard as BMP
    with io.BytesIO() as output:
        image.save(output, format="BMP")
        data = output.getvalue()[14:]  # Skip BMP header
    
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()
    
    doc.close()
    os.remove(temp_pdf)

def calculate_pages():
    """Calculate and display the total number of pages based on inputs."""
    global pages
    file_path = excel_path.get()
    sheet_name = sheet_name_var.get()
    range_address = range_address_var.get()
    orientation = orientation_var.get().lower()
    
    if not file_path or not sheet_name or not range_address or not orientation:
        messagebox.showerror("Error", "Please provide all inputs (Excel file, sheet name, cell range, orientation).")
        return
    
    try:
        temp_file_path = create_temp_sheet_copy(file_path, sheet_name)
        pages = split_range_into_pages(temp_file_path, sheet_name, range_address, orientation)
        total_pages_var.set(f"Total Pages: {len(pages)}")
        status_var.set(f"Calculated {len(pages)} pages.")
        os.remove(temp_file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to calculate pages: {str(e)}")
        status_var.set("Ready")

def create_temp_sheet_copy(file_path, sheet_name):
    """Create a temporary copy of the specified sheet in a new workbook."""
    app = xw.App(visible=False)
    original_wb = app.books.open(file_path)
    temp_wb = app.books.add()
    original_sheet = original_wb.sheets[sheet_name]
    original_sheet.copy(after=temp_wb.sheets[0])
    temp_wb.sheets[0].delete()
    temp_file_path = os.path.join(tempfile.gettempdir(), "temp_excel_sheet.xlsx")
    temp_wb.save(temp_file_path)
    original_wb.close()
    temp_wb.close()
    app.quit()
    return temp_file_path

def capture_and_copy():
    """Capture the selected page, crop it, and copy it to the clipboard."""
    global pages
    if not pages:
        messagebox.showerror("Error", "Please calculate pages first.")
        return
    
    page_num_str = page_num_var.get()
    if not page_num_str.isdigit():
        messagebox.showerror("Error", "Please enter a valid page number.")
        return
    
    page_num = int(page_num_str)
    if page_num < 1 or page_num > len(pages):
        messagebox.showerror("Error", f"Page number must be between 1 and {len(pages)}.")
        return
    
    range_address = pages[page_num - 1]
    quality = quality_var.get()
    dpi = QUALITY_TO_DPI[quality]
    file_path = excel_path.get()
    sheet_name = sheet_name_var.get()
    orientation = orientation_var.get().lower()
    
    # Validate crop ratio
    try:
        crop_ratio = float(crop_ratio_var.get())
        if not 0 < crop_ratio <= 1:
            raise ValueError
    except ValueError:
        messagebox.showerror("Error", "Crop Height Ratio must be a number between 0 and 1 (e.g., 0.77).")
        return
    
    try:
        status_var.set("Processing...")
        root.update_idletasks()
        
        temp_file_path = create_temp_sheet_copy(file_path, sheet_name)
        capture_page(temp_file_path, sheet_name, range_address, orientation, dpi, crop_ratio)
        os.remove(temp_file_path)
        
        status_var.set(f"Page {page_num} copied to clipboard!")
        messagebox.showinfo("Success", f"Page {page_num} has been copied to the clipboard!")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to capture page: {str(e)}")
        status_var.set("Ready")

def browse_excel():
    """Open a file dialog to select an Excel file and populate sheet names."""
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
    if file_path:
        excel_path.set(file_path)
        try:
            sheet_names = get_sheet_names(file_path)
            sheet_name_dropdown['values'] = sheet_names
            if sheet_names:
                sheet_name_var.set(sheet_names[0])
            else:
                sheet_name_var.set("")
            update_calculate_button_state()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet names: {str(e)}")

def update_calculate_button_state():
    """Enable Calculate Pages button only when all inputs are provided."""
    if excel_path.get() and sheet_name_var.get() and range_address_var.get() and orientation_var.get():
        calculate_button.config(state="normal")
    else:
        calculate_button.config(state="disabled")

def show_help():
    """Display help instructions."""
    messagebox.showinfo("Help", """How to Use:
1. Select an Excel file using 'Browse'. The sheet names will be loaded automatically.
2. Choose the sheet name from the dropdown.
3. Enter the cell range (e.g., B2:AD88).
4. Choose the orientation (Portrait or Landscape).
5. Click 'Calculate Pages' to see the total number of pages.
6. Enter the page number you want to capture.
7. Enter the Crop Height Ratio (0-1) to specify how much of the top part to retain (e.g., 0.77 for portrait, 0.785 for landscape).
8. Choose the quality level.
9. Click 'Capture and Copy' to copy the page image to your clipboard.

Quality Levels:
- Low Quality: 100 DPI (less detailed)
- Medium Quality: 300 DPI (balanced)
- High Quality: 600 DPI (highly detailed)

Crop Height Ratio:
- A value between 0 and 1 (e.g., 0.77 retains the top 77% of the image, cropping the bottom 23%).
- Default is 0.77. At 1.0, there is no cropping.
- Example: For portrait, use 0.771 (6.94/9); for landscape, use 0.785 (5/6.37).""")

def close_window():
    """Gracefully close the GUI."""
    root.destroy()

# Set up GUI
root = tk.Tk()
root.title("Excel Page Screenshot Tool")
root.geometry("600x400")
root.resizable(False, False)
root.protocol("WM_DELETE_WINDOW", close_window)

# Excel file selection
excel_path = tk.StringVar()
excel_path.trace("w", lambda *args: update_calculate_button_state())
tk.Label(root, text="Excel File:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=excel_path, width=40, state="readonly").grid(row=0, column=1, columnspan=2, padx=5, pady=10)
tk.Button(root, text="Browse", command=browse_excel).grid(row=0, column=3, padx=5, pady=10)

# Sheet name dropdown
sheet_name_var = tk.StringVar()
sheet_name_var.trace("w", lambda *args: update_calculate_button_state())
tk.Label(root, text="Sheet Name:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
sheet_name_dropdown = ttk.Combobox(root, textvariable=sheet_name_var, state="readonly")
sheet_name_dropdown.grid(row=1, column=1, columnspan=2, padx=5, pady=10)

# Cell range
range_address_var = tk.StringVar()
range_address_var.trace("w", lambda *args: update_calculate_button_state())
tk.Label(root, text="Cell Range (e.g., B2:AD88):").grid(row=2, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=range_address_var, width=40).grid(row=2, column=1, columnspan=2, padx=5, pady=10)

# Orientation
orientation_var = tk.StringVar(value="Landscape")
orientation_var.trace("w", lambda *args: update_calculate_button_state())
tk.Label(root, text="Orientation:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
orientation_dropdown = ttk.Combobox(root, textvariable=orientation_var, values=["Portrait", "Landscape"], state="readonly")
orientation_dropdown.grid(row=3, column=1, padx=5, pady=10)

# Calculate pages button
calculate_button = tk.Button(root, text="Calculate Pages", command=calculate_pages, state="disabled")
calculate_button.grid(row=4, column=1, pady=10)

# Total pages label
total_pages_var = tk.StringVar(value="Total Pages: N/A")
tk.Label(root, textvariable=total_pages_var).grid(row=4, column=2, padx=10, pady=10)

# Page number
page_num_var = tk.StringVar()
tk.Label(root, text="Page Number:").grid(row=5, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=page_num_var, width=10).grid(row=5, column=1, padx=5, pady=10)

# Crop height ratio
crop_ratio_var = tk.DoubleVar(value=0.75)
tk.Label(root, text="Crop Height Ratio (0-1):").grid(row=6, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=crop_ratio_var, width=10).grid(row=6, column=1, padx=5, pady=10)

# Quality
quality_var = tk.StringVar(value="Medium Quality")
tk.Label(root, text="Quality:").grid(row=7, column=0, padx=10, pady=10, sticky="e")
quality_dropdown = ttk.Combobox(root, textvariable=quality_var, values=["Low Quality", "Medium Quality", "High Quality"], state="readonly")
quality_dropdown.grid(row=7, column=1, padx=5, pady=10)

# Capture button
tk.Button(root, text="Capture and Copy", command=capture_and_copy).grid(row=8, column=1, pady=10)

# Status label
status_var = tk.StringVar(value="Ready")
tk.Label(root, textvariable=status_var).grid(row=9, column=0, columnspan=4, pady=10)

# Help button
tk.Button(root, text="Help", command=show_help).grid(row=10, column=1, pady=10)

# Start the GUI
root.mainloop()