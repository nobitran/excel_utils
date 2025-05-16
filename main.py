import openpyxl
from utils import get_col_index, compare_data, print_pretty_differences, export_differences_to_txt
import os

def read_excel_cells(file_path, list_keys, min_row=7):
    """
    Reads all cells from the specified sheet in an Excel (.xlsx) file.

    :param file_path: Path to the Excel file.
    :param sheet_name: Name of the sheet to read.
    :return: A list of lists containing the cell values.
    """
    data = {}
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet_names = workbook.sheetnames
        for sheet in sheet_names:
            sheet_data = workbook[sheet]
            sheet_result = []

            for row in sheet_data.iter_rows(values_only=True, min_row=min_row):
                user = {item["key"]: row[get_col_index(item["value"])] for item in list_keys}
                sheet_result.append(user)
            data[sheet] = sheet_result

        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None
    
def compare_excel_files(file1, file2, list_keys):
    list_sheets1 = read_excel_cells(file1, list_keys)
    list_sheets2 = read_excel_cells(file2, list_keys)
    
    # Compare the two data objects
    differences = compare_data(list_sheets1, list_sheets2)
    
    print(f"Differences for folder {folder}:")
    print_pretty_differences(differences)
    
    # Export the differences to a .txt file
    output_file = f"outputs/{folder}.txt"
    export_differences_to_txt(differences, output_file)
    
    
if __name__ == "__main__":
    base_path_class2 = "files/2"
    
    list_keys = [
        {"key": "order", "value": "A"},
        {"key": "name", "value": "C"},
        {"key": "dob", "value": "D"},
        {"key": "genre", "value": "E"},
        {"key": "ĐĐGtx_1", "value": "F"},
        {"key": "ĐĐGtx_2", "value": "G"},
        {"key": "ĐĐGtx_3", "value": "H"},
        {"key": "ĐĐGtx_4", "value": "I"},
        {"key": "ĐĐGgk", "value": "K"},
        {"key": "ĐĐGck", "value": "L"},
        {"key": "ĐTBmhk1", "value": "M"},
        {"key": "ĐTBmhk2", "value": "N"},
        {"key": "ĐTBmcn", "value": "O"},
    ]
    for folder in os.listdir(base_path_class2):
        folder_path = os.path.join(base_path_class2, folder)
        if os.path.isdir(folder_path):
            files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
            if len(files) >= 2:
                file1 = os.path.join(folder_path, files[0])
                file2 = os.path.join(folder_path, files[1])
                compare_excel_files(file1, file2, list_keys)
               