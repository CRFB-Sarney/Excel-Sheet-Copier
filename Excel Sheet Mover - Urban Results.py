# This program copies a sheet from one workbook to another, preserving formatting and values, but not formulas.
# It works with Excel 2010 and later. It uses openpyxl library to handle Excel files.

import openpyxl
from openpyxl import load_workbook

import tkinter as tk
from tkinter import filedialog, messagebox
import copy

from pathlib import Path

from datetime import datetime

# Function to copy a sheet from source workbook to destination workbook
def copy_sheet(source_file, out_file, source_sheet_name, out_sheet_name):
    out_file = Path(out_file)

    # Load the source workbook once for formatting and once for cached values.
    source_wb = load_workbook(source_file, data_only=False)
    source_values_wb = load_workbook(source_file, data_only=True)
    if source_sheet_name not in source_wb.sheetnames:
        messagebox.showerror("Error", f"Sheet '{source_sheet_name}' not found in source workbook.")
        return
    
    source_sheet = source_wb[source_sheet_name]
    source_values_sheet = source_values_wb[source_sheet_name]

    # Load or create the destination workbook
    if out_file.exists():
        out_wb = load_workbook(out_file)
    else:
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)  # Remove the default sheet

    # Create a new sheet in the destination workbook with the same name
    if out_sheet_name in out_wb.sheetnames:
        messagebox.showerror("Error", f"Sheet '{out_sheet_name}' already exists in destination workbook.")
        return
    
    out_sheet = out_wb.create_sheet(title=out_sheet_name)

    # Copy the sheet tab color, if present.
    try:
        if source_sheet.sheet_properties.tabColor is not None:
            out_sheet.sheet_properties.tabColor = copy.copy(source_sheet.sheet_properties.tabColor)
    except Exception:
        pass

    # Copy cell values and formatting from source to destination
    for row in source_sheet.iter_rows():
        for cell in row:
            value = source_values_sheet.cell(row=cell.row, column=cell.column).value
            new_cell = out_sheet.cell(row=cell.row, column=cell.column, value=value)
            if cell.has_style:
                new_cell.font = copy.copy(cell.font)
                new_cell.border = copy.copy(cell.border)
                new_cell.fill = copy.copy(cell.fill)
                new_cell.number_format = copy.copy(cell.number_format)
                new_cell.protection = copy.copy(cell.protection)
                new_cell.alignment = copy.copy(cell.alignment)

    # Copy merged cell ranges, such as A1:C1.
    for merged_range in source_sheet.merged_cells.ranges:
        out_sheet.merge_cells(str(merged_range))

    # copy column widths
    try:
        for col, dim in source_sheet.column_dimensions.items():
            if dim and dim.width:
                out_sheet.column_dimensions[col].width = dim.width
    except Exception:
        pass

    # copy row heights
    try:
        for idx, dim in source_sheet.row_dimensions.items():
            if dim and dim.height:
                out_sheet.row_dimensions[idx].height = dim.height
    except Exception:
        pass
    
    # Save the destination workbook
    out_wb.save(out_file)

    print(f"At {datetime.now().strftime('%H:%M:%S')}: Copied '{source_sheet_name}' to '{out_sheet_name}'.")

# GUI setup to allow user to select source and destination files and specify the sheet name
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Ask user to select source workbook
    source_file = filedialog.askopenfilename(title="Select Source Workbook", filetypes=[("Excel files", "*.xlsx")])
    if not source_file:
        messagebox.showerror("Error", "No source workbook selected.")
        return

    # Ask user to select destination workbook
    out_file = filedialog.askopenfilename(title="Select Destination Workbook", filetypes=[("Excel files", "*.xlsx")])
    if not out_file:
        messagebox.showerror("Error", "No destination workbook selected.")
        return

    # Ask user to specify destination sheet name prefix
    out_sheet_prefix = tk.simpledialog.askstring("Input", "Enter the prefix for the destination sheet names:")
    if not out_sheet_prefix:
        messagebox.showerror("Error", "No destination sheet name prefix provided.")
        return

    # Call the function to copy the sheets
    print(f"At {datetime.now().strftime('%H:%M:%S')}: began copying sheets.")
    copy_sheet(source_file, out_file, "10-year budget", f"{out_sheet_prefix} 10yr")
    copy_sheet(source_file, out_file, "OCACT estimate", f"{out_sheet_prefix} TF %TP")
    copy_sheet(source_file, out_file, "$SSBSSIbyYearAndQuintile", f"{out_sheet_prefix} SSBSSI")
    copy_sheet(source_file, out_file, "poverty compare", f"{out_sheet_prefix} Pov")   
    print(f"Done at {datetime.now().strftime('%H:%M:%S')}!", f"Sheets from '{source_file}' copied to '{out_file}'.")

#Run program
main()