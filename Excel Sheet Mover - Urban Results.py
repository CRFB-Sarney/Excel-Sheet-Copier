# This program copies a sheet from one workbook to another, preserving formatting and values, but not formulas.
# It works with Excel 2010 and later. It uses openpyxl library to handle Excel files.

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries, get_column_letter

import tkinter as tk
from tkinter import filedialog, messagebox
import copy

from pathlib import Path

from datetime import datetime

# Function to copy a sheet from source workbook to destination workbook
def copy_sheet(source_file, source_sheet_name, source_range, out_file, out_sheet_name):
    out_file = Path(out_file)

    # Load the source workbook once for formatting and once for cached values.
    source_wb = load_workbook(source_file, data_only=False)
    source_values_wb = load_workbook(source_file, data_only=True)
    if source_sheet_name not in source_wb.sheetnames:
        print("Error", f"Sheet '{source_sheet_name}' not found in source workbook.")
        return
    
    source_sheet = source_wb[source_sheet_name]
    source_values_sheet = source_values_wb[source_sheet_name]
    source_cellblock = source_sheet[source_range]
    source_start = source_cellblock[0][0]
    target_start_row = 1 # Row 1
    target_start_col = 1 # Column A
    min_col, min_row, max_col, max_row = range_boundaries(source_range)

    # Load or create the destination workbook
    if out_file.exists():
        out_wb = load_workbook(out_file)
    else:
        out_wb = openpyxl.Workbook()
        out_wb.remove(out_wb.active)  # Remove the default sheet

    # Create a new sheet in the destination workbook with the same name
    if out_sheet_name in out_wb.sheetnames:
        print("Note", f"Sheet '{out_sheet_name}' already exists and will be overwritten.")
        return
    
    out_sheet = out_wb.create_sheet(title=out_sheet_name)

    # Copy the sheet tab color, if present.
    try:
        if source_sheet.sheet_properties.tabColor is not None:
            out_sheet.sheet_properties.tabColor = copy.copy(source_sheet.sheet_properties.tabColor)
    except Exception:
        pass

    # Copy cell values and formatting from source to destination
    for source_row in source_cellblock:
        for source_cell in source_row:
            target_row = source_cell.row - source_start.row + 1
            target_col = source_cell.column - source_start.column + 1
            value = source_values_sheet.cell(row=source_cell.row, column=source_cell.column).value
            target_cell = out_sheet.cell(row=target_row, column=target_col, value=value)
            if source_cell.has_style:
                target_cell.font = copy.copy(source_cell.font)
                target_cell.border = copy.copy(source_cell.border)
                target_cell.fill = copy.copy(source_cell.fill)
                target_cell.number_format = copy.copy(source_cell.number_format)
                target_cell.protection = copy.copy(source_cell.protection)
                target_cell.alignment = copy.copy(source_cell.alignment)

    # Copy merged cell ranges, such as A1:C1.
    for merged_range in source_sheet.merged_cells.ranges:
        merged_min_col, merged_min_row, merged_max_col, merged_max_row = \
             range_boundaries(str(merged_range))
        if (
            merged_min_row >= source_start.row 
            and merged_max_row <= source_start.row + len(source_cellblock) - 1 
            and merged_min_col >= source_start.column 
            and merged_max_col <= source_start.column + len(source_cellblock[0]) - 1
        ):
            target_min_col = target_start_col + (merged_min_col - source_start.column)
            target_min_row = target_start_row + (merged_min_row - source_start.row)
            target_max_col = target_start_col + (merged_max_col - source_start.column)
            target_max_row = target_start_row + (merged_max_row - source_start.row)
            
            out_sheet.merge_cells(
                f"{get_column_letter(target_min_col)}{target_min_row}:"
                f"{get_column_letter(target_max_col)}{target_max_row}"
                )

    # copy column widths from source to corresponding target
    try:
        for source_col in range(min_col, max_col + 1):
            source_col_letter = get_column_letter(source_col)
            target_col_letter = get_column_letter(target_start_col + (source_col - min_col))
            if source_col_letter in source_sheet.column_dimensions:
                source_dim = source_sheet.column_dimensions[source_col_letter]
                target_dim = out_sheet.column_dimensions[target_col_letter]
                target_dim.width = source_dim.width
    except Exception:
        pass

    # copy row heights
    try:
        for source_row in range(min_row, max_row + 1):
            source_row_letter = get_column_letter(source_row)
            target_row_letter = get_column_letter(target_start_row + (source_row - min_row))
            if source_row_letter in source_sheet.row_dimensions:
                source_dim = source_sheet.row_dimensions[source_row_letter]
                target_dim = out_sheet.row_dimensions[target_row_letter]
                target_dim.height = source_dim.height
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
    copy_sheet(source_file, "10-year budget", "AQ3:AS19", out_file, f"{out_sheet_prefix} 10yr")
    copy_sheet(source_file, "OCACT estimate", "A6:Y115", out_file, f"{out_sheet_prefix} TF %TP")
    copy_sheet(source_file, "$SSBSSIbyYearAndQuintile", "T617:AK916", out_file, f"{out_sheet_prefix} SSBSSI")
    copy_sheet(source_file, "poverty compare", "A1:I150", out_file, f"{out_sheet_prefix} Pov")
    print(f"Done at {datetime.now().strftime('%H:%M:%S')}!", f"Sheets from '{source_file}' copied to '{out_file}'.")

#Run program
main()