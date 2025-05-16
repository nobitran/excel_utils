

def get_col_index(col_letter):
    """
    Converts an Excel column letter to a zero-based index.

    :param col_letter: The column letter (e.g., 'A', 'B', 'C').
    :return: The zero-based column index.
    """
    return ord(col_letter.upper()) - ord('A')


def compare_data(data1, data2):
    differences = {}
    for subject, records1 in data1.items():
        records2 = data2.get(subject, [])
        for i, record1 in enumerate(records1):
            if i < len(records2):
                record2 = records2[i]
                for key, value1 in record1.items():
                    value2 = record2.get(key)
                    try:
                        value1 = float(value1)
                        value2 = float(value2)
                    except (ValueError, TypeError):
                        pass

                    if value1 != value2:
                        if subject not in differences:
                            differences[subject] = {}
                        name = record1.get("name")
                        if name:
                            if name not in differences[subject]:
                                differences[subject][name] = []
                            differences[subject][name].append({
                                "field": key,
                                "value1": value1,
                                "value2": value2
                            })
    return differences

def print_pretty_differences(differences):
    for subject, students in differences.items():
        print(f"- {subject} có {len(students)} học sinh có sự khác biệt")
        for student, fields in students.items():
            field_differences = "; ".join(
                [f"{field['field']}: ({field['value1']} != {field['value2']})" for field in fields]
            )
            print(f"  + {student}: {field_differences}")    
    
def export_differences_to_txt(differences, output_file):
    """
    Exports the differences to a .txt file.

    :param differences: The differences dictionary.
    :param output_file: Path to the output .txt file.
    """
    try:
        with open(output_file, "w") as file:
            for subject, students in differences.items():
                file.write(f"- {subject} có {len(students)} học sinh có sự khác biệt\n")
                for student, fields in students.items():
                    field_differences = "; ".join(
                        [f"{field['field']}: ({field['value1']} != {field['value2']})" for field in fields]
                    )
                    file.write(f"  + {student}: {field_differences}\n")
        print(f"Differences exported to {output_file}")
    except Exception as e:
        print(f"Error exporting differences to file: {e}")

    