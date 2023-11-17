import tkinter as tk
from tkinter import filedialog, ttk
import csv
import pandas as pd
import os
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import xlwings as xw


def csv_to_dataframe(file_path):
    return pd.read_csv(file_path)


def read_csv_columns(file_path):
    print(f"Reading columns from: {file_path}")
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        return next(reader, [])  # Return the first row (headers)


def update_common_columns(file1_entry, file2_entry, combobox):
    file1 = file1_entry.get()
    file2 = file2_entry.get()

    if file1 and file2:
        print("Both files selected. Finding common columns...")
        columns_file1 = read_csv_columns(file1)
        columns_file2 = read_csv_columns(file2)

        # Finding common columns
        common_columns = [col for col in columns_file1 if col in columns_file2]
        print(f"Common columns found: {common_columns}")

        # Updating the dropdown menu
        combobox['values'] = common_columns

        if common_columns:
            combobox.current(0)
        else:
            print("No common columns found.")
    else:
        print("Both files are not yet selected.")


def open_file(entry, combobox, other_entry):
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)
        print(f"File selected: {file_path}")
        update_common_columns(file1_entry, file2_entry, combobox)
    else:
        print("File selection cancelled.")


def compare_csvs(file1, file2, key1_index, key2_index):
    print(f"Comparing files: {file1} and {file2}")
    csv1_dict = read_csv_to_dict(file1, key1_index)
    csv2_dict = read_csv_to_dict(file2, key2_index)

    all_keys = set(csv1_dict.keys()) | set(csv2_dict.keys())
    differences = []

    for key in all_keys:
        row1 = csv1_dict.get(key)
        row2 = csv2_dict.get(key)

        if row1 != row2:
            differences.append((key, row1, row2))

    return differences


def read_csv_to_dict(file_path, key_column):
    print(f"Reading file into dictionary: {file_path}")
    data_dict = {}
    with open(file_path, newline='', encoding='utf-8') as f:
        reader = csv.reader(f)
        for row in reader:
            key = row[key_column]
            data_dict[key] = row
    return data_dict


# Modify the report_differences function to call save_to_excel
def report_differences(differences, file1, file2, output_path):
    if not differences:
        print("No differences found.")
    else:
        save_to_excel(file1, file2, differences, output_path)


def compare_and_report():
    file1 = file1_entry.get()
    file2 = file2_entry.get()
    common_key = key_combobox.get()
    output_folder = output_folder_entry.get()

    # Check if both files and a key are selected
    if file1 and file2 and common_key and output_folder:
        print(f"Comparing files: {file1}, {file2} using key: {common_key}")
        output_path = os.path.join(output_folder, "comparison_output.xlsx")
        # Call the save_to_excel function with necessary arguments
        save_to_excel(file1, file2, output_path, common_key)
        print(f"Excel file with comparison saved at: {output_path}")
    else:
        print("Please select both files, a common key, and an output folder")


def save_to_excel(original_file, comparison_file, output_path, common_key):
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Original file data
        original_df = csv_to_dataframe(original_file)
        original_df.to_excel(writer, sheet_name="Original file", index=False)

        # Comparison file data
        comparison_df = csv_to_dataframe(comparison_file)
        comparison_df.to_excel(writer, sheet_name="Export file", index=False)

        # Format the 'Compare' sheet as specified
        format_compare_sheet(writer, original_df, comparison_df, common_key)

        print(f"Excel file saved: {output_path}")

    output_folder = output_folder_entry.get()
    output_path = os.path.join(output_folder, "comparison_output.xlsx")
    auto_fill_formula(last_col=get_column_letter(len(original_df.columns) * 3 + 1), last_row=len(original_df)+2, path=output_path)


def format_compare_sheet(writer, original_df, comparison_df, common_key):
    # Create the header with the column names spread across every third column, starting from column B
    header_columns = [''] + sum([[col, '', ''] for col in original_df.columns], [])

    # Initialize the DataFrame with an empty row for the main headers
    compare_df = pd.DataFrame(columns=header_columns)
    compare_df.loc[0] = [''] * len(header_columns)  # Empty row for main headers

    # Prepare the sub-headers "Original", "Export", "Compare" for each column in original_df
    sub_headers = ['Key']
    for _ in original_df.columns:
        sub_headers.extend(['Original', 'Export', 'Compare'])

    # Insert the sub-headers into the second row
    compare_df.loc[0] = sub_headers

    # Add the key data to the DataFrame starting from the third row
    key_data = original_df[common_key].tolist()
    for i, key in enumerate(key_data):
        # Starting index is 2 because we already have headers and sub-headers
        compare_df.loc[i + 1] = [key] + [''] * (len(header_columns) - 1)

    # Writing the 'Compare' DataFrame to the Excel writer
    compare_df.to_excel(writer, sheet_name="Compare", index=False)

    # Access the workbook and the specific sheet to apply formatting
    workbook = writer.book
    worksheet = writer.sheets['Compare']

    # Apply left alignment to the first row
    for cell in worksheet['1:1']:  # Accessing the first row
        cell.alignment = Alignment(horizontal='left')

    set_original_values_formula(worksheet, original_df, len(compare_df.columns), comparison_df)



def choose_output_folder(entry):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry.delete(0, tk.END)
        entry.insert(0, folder_path)
        print(f"Output folder selected: {folder_path}")


def set_original_values_formula(worksheet, original_df, num_cols, comparison_df):
    max_row = len(original_df) + 2  # Adjust based on your data
    max_col = get_column_letter(num_cols)
    original_range = range(2, len(original_df.columns) * 3, 3)
    export_range = range(3,  (len(original_df.columns)+1) * 3, 3)
    compare_range = range(4,  (len(original_df.columns)+1) * 3, 3)
    compare_length = len(comparison_df) + 1
    original_length = len(original_df) + 1
    key_column_name = key_combobox.get()
    original_key_column = get_column_letter(original_df.columns.get_loc(key_column_name) + 1)
    export_key_column = get_column_letter(comparison_df.columns.get_loc(key_column_name) + 1)
    print(f"column in Original tab: {original_key_column}\nColumn in export tab: {export_key_column}")


    # TODO: How to make sure that Column B is the Key (or should I get any indication on which column?)
    # TODO: Insert the export formula and compare formula
    for idx in original_range: # Original formula insertion
        col_letter = get_column_letter(idx)
        print(f"column latter:{col_letter}")
        cell = worksheet.cell(row=3, column=idx)
        formula = f"=HLOOKUP(${col_letter}$1,'Original file'!$1:${original_length},MATCH($A3,'Original file'!${original_key_column}:${original_key_column},0),FALSE)"
        cell.value = formula

    for idx in export_range: # Export formula insertion
        col_letter = get_column_letter(idx - 1)
        print(f"column latter:{col_letter}")
        cell = worksheet.cell(row=3, column=idx)
        formula = f"=HLOOKUP(${col_letter}$1,'Export file'!$1:${compare_length},MATCH($A3,'Export file'!${export_key_column}:${export_key_column},0),FALSE)"
        cell.value = formula

    for idx in compare_range: # Export formula insertion
        og_col = get_column_letter(idx - 2)
        export_col = get_column_letter(idx - 1)
        #col_letter = get_column_letter(idx - 2)
        #print(f"column latter:{col_letter}")
        cell = worksheet.cell(row=3, column=idx)
        formula = f'=IF(OR(ISNA({og_col}3),ISNA({export_col}3)),"Error",IF({og_col}3={export_col}3,"OK","Error"))'
        # =IF(OR(ISNA(B3),ISNA(C3)),"Error",IF(B3=C3,"OK","Error"))
        cell.value = formula

def auto_fill_formula(last_col, last_row, path):
    # Open an existing workbook
    wb = xw.Book(path)  # Replace with your file path
    sheet = wb.sheets['Compare']  # Replace with your sheet name

    # Select the cell with the formula
    cell_with_formula = sheet.range(f"B3:{last_col}3")

    # Define the range to auto-fill
    # For example, auto-filling from A2 to A10
    fill_range = sheet.range(f"B3:{last_col}{last_row}")

    # Use the auto-fill
    cell_with_formula.api.AutoFill(fill_range.api, xw.constants.AutoFillType.xlFillDefault)

    # Save and close the workbook
    wb.save()
    wb.close()


# GUI setup
root = tk.Tk()
root.title("CSV Comparison Tool")

# File selection for File 1
file1_frame = tk.Frame(root)
file1_frame.pack(padx=10, pady=5)
file1_entry = tk.Entry(file1_frame, width=50)
file1_entry.pack(side=tk.LEFT)
file1_button = tk.Button(file1_frame, text="Select File 1", command=lambda: open_file(file1_entry, key_combobox, file2_entry))
file1_button.pack(side=tk.LEFT)

# File selection for File 2
file2_frame = tk.Frame(root)
file2_frame.pack(padx=10, pady=5)
file2_entry = tk.Entry(file2_frame, width=50)
file2_entry.pack(side=tk.LEFT)
file2_button = tk.Button(file2_frame, text="Select File 2", command=lambda: open_file(file2_entry, key_combobox, file1_entry))
file2_button.pack(side=tk.LEFT)

# Key selection (common columns)
key_frame = tk.Frame(root)
key_frame.pack(padx=10, pady=5)
key_combobox = ttk.Combobox(key_frame, width=15)
key_combobox.pack(side=tk.LEFT)

# Output folder selection
output_folder_frame = tk.Frame(root)
output_folder_frame.pack(padx=10, pady=5)
output_folder_entry = tk.Entry(output_folder_frame, width=50)
output_folder_entry.pack(side=tk.LEFT)
output_folder_button = tk.Button(output_folder_frame, text="Select Output Folder", command=lambda: choose_output_folder(output_folder_entry))
output_folder_button.pack(side=tk.LEFT)

# 'Compare Files' button
compare_button = tk.Button(root, text="Compare Files", command=compare_and_report)
compare_button.pack(pady=10)

# Start the GUI
root.mainloop()
