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

def copy_multiple_sheets():
    # Call the function to copy the sheets
    print(f"At {datetime.now().strftime('%H:%M:%S')}: began copying sheets.")
    # Load the source workbook once for formatting and once for cached values.

    # Ask user to select source workbook
    source_file = filedialog.askopenfilename(title="Select Source Workbook", filetypes=[("Excel files", "*.xlsx")])
    if not source_file:
        messagebox.showerror("Error", "No source workbook selected.")
        return

    # Ask user to select destination workbook
    target_file = filedialog.askopenfilename(title="Select Destination Workbook", filetypes=[("Excel files", "*.xlsx")])
    if not target_file:
        messagebox.showerror("Error", "No destination workbook selected.")
        return

    # Ask user to specify destination sheet name prefix
    target_sheet_prefix = tk.simpledialog.askstring("Input", "Enter the prefix for the destination sheet names:")
    if not target_sheet_prefix:
        messagebox.showerror("Error", "No destination sheet name prefix provided.")
        return

    source_styles_wb = load_workbook(source_file, data_only=False)
    source_values_wb = load_workbook(source_file, data_only=True)
    target_file = Path(target_file)
    # Load or create the destination workbook
    if target_file.exists():
        target_wb = load_workbook(target_file)
    else:
        target_wb = openpyxl.Workbook()
        target_wb.remove(target_wb.active)  # Remove the default sheet

    copy_sheet(source_styles_wb, source_values_wb, "10-year budget", "AQ3:AS19", target_wb, f"{target_sheet_prefix} 10yr")
    copy_sheet(source_styles_wb, source_values_wb, "OCACT estimate", "A6:Y115", target_wb, f"{target_sheet_prefix} TF %TP")
    copy_sheet(source_styles_wb, source_values_wb, "$SSBSSIbyYearAndQuintile", "T617:AK916", target_wb, f"{target_sheet_prefix} SSBSSI")
    copy_sheet(source_styles_wb, source_values_wb, "poverty compare", "A1:I150", target_wb, f"{target_sheet_prefix} Pov")

    # Save the destination workbook
    target_wb.save(target_file)

    print(f"Done at {datetime.now().strftime('%H:%M:%S')}!", f"Sheets from '{source_file}' copied to '{target_file}'.")

# Function to copy a sheet from a predefined source workbook to a predefined destination workbook
def copy_sheet(source_styles_wb, source_values_wb, source_sheet_name, source_range, target_wb, target_sheet_name):

    if source_sheet_name not in source_styles_wb.sheetnames:
        print("Error", f"Sheet '{source_sheet_name}' not found in source workbook.")
        return
      
    source_styles_sheet = source_styles_wb[source_sheet_name]
    source_values_sheet = source_values_wb[source_sheet_name]
    source_cellblock = source_styles_sheet[source_range]
    source_start = source_cellblock[0][0]
    target_row_start = 1 # Row 1
    target_col_start = 1 # Column A
    min_col, min_row, max_col, max_row = range_boundaries(source_range)

    # Create a new sheet in the destination workbook with the same name
    if target_sheet_name in target_wb.sheetnames:
        print("Note", f"Sheet '{target_sheet_name}' already exists and will be overwritten.")
        return
    
    target_sheet = target_wb.create_sheet(title=target_sheet_name)

    # Copy the cell values in the block
    copy_cellblock(source_values_sheet, source_range, target_sheet, target_row_start, target_col_start)


    # Copy the sheet tab color, if present.
    try:
        if source_styles_sheet.sheet_properties.tabColor is not None:
            target_sheet.sheet_properties.tabColor = copy.copy(source_styles_sheet.sheet_properties.tabColor)
    except Exception:
        pass

    # Copy merged cell ranges, such as A1:C1.
    for merged_range in source_styles_sheet.merged_cells.ranges:
        merged_min_col, merged_min_row, merged_max_col, merged_max_row = \
             range_boundaries(str(merged_range))
        if (
            merged_min_row >= source_start.row 
            and merged_max_row <= source_start.row + len(source_cellblock) - 1 
            and merged_min_col >= source_start.column 
            and merged_max_col <= source_start.column + len(source_cellblock[0]) - 1
        ):
            target_min_col = target_col_start + (merged_min_col - source_start.column)
            target_min_row = target_row_start + (merged_min_row - source_start.row)
            target_max_col = target_col_start + (merged_max_col - source_start.column)
            target_max_row = target_row_start + (merged_max_row - source_start.row)
            
            target_sheet.merge_cells(
                f"{get_column_letter(target_min_col)}{target_min_row}:"
                f"{get_column_letter(target_max_col)}{target_max_row}"
                )

    # copy column widths from source to corresponding target
    try:
        for source_col in range(min_col, max_col + 1):
            source_col_letter = get_column_letter(source_col)
            target_col_letter = get_column_letter(target_col_start + (source_col - min_col))
            if source_col_letter in source_styles_sheet.column_dimensions:
                source_dim = source_styles_sheet.column_dimensions[source_col_letter]
                target_dim = target_sheet.column_dimensions[target_col_letter]
                target_dim.width = source_dim.width
    except Exception:
        pass

    # copy row heights
    try:
        for source_row in range(min_row, max_row + 1):
            source_row_letter = get_column_letter(source_row)
            target_row_letter = get_column_letter(target_row_start + (source_row - min_row))
            if source_row_letter in source_styles_sheet.row_dimensions:
                source_dim = source_styles_sheet.row_dimensions[source_row_letter]
                target_dim = target_sheet.row_dimensions[target_row_letter]
                target_dim.height = source_dim.height
    except Exception:
        pass
    
    print(f"At {datetime.now().strftime('%H:%M:%S')}: Copied '{source_sheet_name}' to '{target_sheet_name}'.")

#This copies a block of cells and their formats and pastes them in a different cell block from the source row and column
def copy_cellblock(source_values_sheet, source_range, target_sheet, target_row_start, target_col_start):
    min_col, min_row, max_col, max_row = range_boundaries(source_range)
    for r_offset, source_row in enumerate(
        source_values_sheet.iter_rows(
            min_row=min_row,
            max_row=max_row,
            min_col=min_col,
            max_col=max_col),
            start=0
        ):
        for col_offset, source_cell in enumerate(source_row, start=0):
            target_cell = target_sheet.cell(row=target_row_start + r_offset, column=target_col_start + col_offset)
            target_cell.value = source_cell.value

            if source_cell.has_style:
                target_cell.font = copy.copy(source_cell.font)
                target_cell.border = copy.copy(source_cell.border)
                target_cell.fill = copy.copy(source_cell.fill)
                target_cell.number_format = copy.copy(source_cell.number_format)
                target_cell.protection = copy.copy(source_cell.protection)
                target_cell.alignment = copy.copy(source_cell.alignment)
  

# Main program - makkkking coppppieees
def making_copies():
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    copy_multiple_sheets()

#Run program
making_copies()