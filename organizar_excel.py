import os
import pandas as pd

def separate_excel_sheets(input_directory):
    """
    Iterates through all Excel files in a directory and saves each sheet as a new Excel file.

    Args:
        input_directory (str): The path to the directory containing the Excel files.
    """
    print(f"Searching for Excel files in: {input_directory}")

    for filename in os.listdir(input_directory):
        if filename.endswith((".xlsx", ".xls")):
            file_path = os.path.join(input_directory, filename)
            print(f"Processing file: {file_path}")

            try:
                xls = pd.ExcelFile(file_path)
                for sheet_name in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet_name)
                    
                    # Sanitize sheet name to create a valid filename by removing illegal characters
                    invalid_chars = '<>:"/\\|?*'
                    sanitized_sheet_name = "".join(c for c in sheet_name if c not in invalid_chars).strip()
                    output_filename = f"{sanitized_sheet_name}.xlsx"
                    output_path = os.path.join(input_directory, output_filename)
                    
                    print(f"  - Saving sheet '{sheet_name}' to '{output_path}'")
                    
                    # Save the sheet to a new Excel file
                    df.to_excel(output_path, index=False)
                print(f"Finished processing {filename}")
            except Exception as e:
                print(f"Could not process file {filename}. Error: {e}")

if __name__ == "__main__":
    # The user wants to process files in the "C:\Users\walyson.ferreira\Desktop\ORSE" directory
    target_directory = r"C:\Users\walyson.ferreira\Desktop\ORSE"
    
    if os.path.isdir(target_directory):
        separate_excel_sheets(target_directory)
        print("\nAll Excel files have been processed.")
    else:
        print(f"Error: The directory '{target_directory}' does not exist.")

# To run this script:
# 1. Make sure you have pandas and openpyxl installed:
#    pip install pandas openpyxl
# 2. Save this code as a Python file (e.g., organizar_excel.py).
# 3. Run the script from your terminal:
#    python organizar_excel.py
