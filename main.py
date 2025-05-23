import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import fitz  # PyMuPDF
from PIL import Image
import win32clipboard
import io
import os
import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def convert_and_copy():
    pdf_file = pdf_path.get()
    page_num = page_entry.get()
    quality = quality_var.get()

    # Input validation
    if not pdf_file:
        messagebox.showerror("Error", "Please select a PDF file.")
        return
    if not page_num.isdigit() or int(page_num) < 1:
        messagebox.showerror("Error", "Please enter a positive integer for the page number.")
        return

    # Map quality to DPI
    quality_map = {
        "Low Quality": 100,
        "Medium Quality": 300,
        "High Quality": 600
    }
    dpi = quality_map[quality]

    try:
        status_var.set("Processing...")
        root.update_idletasks()

        # Open PDF and render page
        doc = fitz.open(pdf_file)
        if int(page_num) > doc.page_count:
            doc.close()
            raise ValueError(f"Page {page_num} does not exist. PDF has {doc.page_count} pages.")

        page = doc.load_page(int(page_num) - 1)  # 0-based index
        zoom = dpi / 72  # PyMuPDF default resolution is 72 DPI
        mat = fitz.Matrix(zoom, zoom)  # Scale matrix for DPI
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img_data = pix.tobytes("png")  # Get image as PNG

        # Convert to PIL Image for clipboard
        image = Image.open(io.BytesIO(img_data))

        # Save image to bytes buffer for clipboard (BMP format)
        with io.BytesIO() as output:
            image.save(output, format="BMP")
            data = output.getvalue()[14:]  # Skip BMP header

        # Copy to clipboard
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
        win32clipboard.CloseClipboard()

        doc.close()
        status_var.set("Image copied to clipboard!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        status_var.set("Ready")

def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        pdf_path.set(file_path)

def show_help():
    messagebox.showinfo("Help", "How to Use:\n"
                        "1. Click 'Browse' to select a PDF file.\n"
                        "2. Enter the page number you want to capture.\n"
                        "3. Choose a quality level from the dropdown.\n"
                        "4. Click 'Convert and Copy' to copy the image to your clipboard.\n\n"
                        "Quality Levels:\n"
                        "- Low Quality: 100 DPI (smaller, less detailed)\n"
                        "- Medium Quality: 300 DPI (balanced)\n"
                        "- High Quality: 600 DPI (larger, more detailed)")

def close_window():
    root.destroy()  # Gracefully close the GUI

# Set up GUI
root = tk.Tk()
root.title("PDF Page to Clipboard")
root.geometry("500x250")
root.resizable(False, False)

# Handle window close event
root.protocol("WM_DELETE_WINDOW", close_window)

# PDF selection
pdf_path = tk.StringVar()
tk.Label(root, text="PDF File:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
tk.Entry(root, textvariable=pdf_path, width=40, state="readonly").grid(row=0, column=1, columnspan=2, padx=5, pady=10)
tk.Button(root, text="Browse", command=browse_pdf).grid(row=0, column=3, padx=5, pady=10)

# Page number
tk.Label(root, text="Page Number:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
page_entry = tk.Entry(root, width=10)
page_entry.grid(row=1, column=1, padx=5, pady=10, sticky="w")

# Quality dropdown
tk.Label(root, text="Quality:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
quality_var = tk.StringVar(value="Medium Quality")
quality_dropdown = ttk.Combobox(root, textvariable=quality_var, values=["Low Quality", "Medium Quality", "High Quality"], state="readonly")
quality_dropdown.grid(row=2, column=1, padx=5, pady=10, sticky="w")

# Convert button
tk.Button(root, text="Convert and Copy", command=convert_and_copy).grid(row=3, column=1, columnspan=2, pady=10)

# Status label
status_var = tk.StringVar(value="Ready")
tk.Label(root, textvariable=status_var).grid(row=4, column=0, columnspan=4, pady=10)

# Help button
tk.Button(root, text="Help", command=show_help).grid(row=5, column=1, columnspan=2, pady=10)

root.mainloop()