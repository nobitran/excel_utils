import openpyxl
from utils import get_col_index, compare_data, print_pretty_differences, export_differences_to_txt
import os
import json
from config import LIST_KEY_FILE1, LIST_KEY_FILE2, MAPPING_FIELD, EXTENSION_FILE1, EXTENSION_FILE2, BASE_FOLDER

def read_excel_cells(file_path, list_keys, min_row=7):
    """
    Reads all cells from the specified sheet in an Excel (.xlsx) file.

    :param file_path: Path to the Excel file.
    :param sheet_name: Name of the sheet to read.
    :return: A list of lists containing the cell values.
    """
    data = {}
    try:
        workbook = openpyxl.load_workbook(file_path, data_only=True)
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
    
def compare_excel_files(file1_data, file2_data, mapping, folder):
    list_sheets1 = read_excel_cells(file1_data["file_path"], file1_data["list_keys"], file1_data["min_row"])
    list_sheets2 = read_excel_cells(file2_data["file_path"], file2_data["list_keys"], file2_data["min_row"])
    # Export the data to JSON files
    os.makedirs(f"json/{folder}", exist_ok=True)
    with open(f"json/{folder}/{file1_data["file_name"]}.json", "w", encoding="utf-8") as f1:
        json.dump(list_sheets1, f1, ensure_ascii=False, indent=4)

    with open(f"json/{folder}/{file2_data["file_name"]}.json", "w", encoding="utf-8") as f2:
        json.dump(list_sheets2, f2, ensure_ascii=False, indent=4)
    
    # Compare the two data objects
    differences = compare_data(list_sheets1, list_sheets2, mapping)
    
    if not differences:
        print(f"No differences found in {folder}.")
    else:
        print_pretty_differences(differences)
    
    # Export the differences to a .txt file
    output_file = f"outputs/{folder}.txt"
    export_differences_to_txt(differences, output_file)
    
    
if __name__ == "__main__":

    for folder in os.listdir(BASE_FOLDER):
        folder_path = os.path.join(BASE_FOLDER, folder)
        if os.path.isdir(folder_path):
            files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx") and not f.startswith("~$")]
            if len(files) >= 2:
                file1 = next((os.path.join(folder_path, f) for f in files if f.endswith(EXTENSION_FILE1)), None)
                file2 = next((os.path.join(folder_path, f) for f in files if f.endswith(EXTENSION_FILE2)), None)
                file_name1 = os.path.splitext(os.path.basename(file1))[0]
                file_name2 = os.path.splitext(os.path.basename(file2))[0]
                if file1 and file2:
                    file1_data = {
                        "list_keys": LIST_KEY_FILE1,
                        "file_path": file1,
                        "min_row": 9,
                        "file_name": file_name1
                    }
                    file2_data = {
                        "list_keys": LIST_KEY_FILE2,
                        "file_path": file2,
                        "min_row": 7,
                        "file_name": file_name2
                    }
                    compare_excel_files(file1_data, file2_data, MAPPING_FIELD, folder)
