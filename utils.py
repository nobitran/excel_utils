# This script contains utility functions for comparing two sets of data, specifically for checking differences in student records across different subjects.

def get_col_index(col_letter):
    return ord(col_letter.upper()) - ord('A')

def get_subject_for_data2(mapping, subject):
    return mapping.get(subject, subject)

def compare_data(data1, data2, mapping):
    differences = {}
    for subject, records1 in data1.items():
        subject2 = get_subject_for_data2(mapping, subject)
        records2 = {record.get("name"): record for record in data2.get(subject2, [])}
        for record1 in records1:
            name = record1.get("name")
            if name in records2:
                record2 = records2[name]
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
                        if name not in differences[subject]:
                            differences[subject][name] = []
                        differences[subject][name].append({
                            "field": key,
                            "value1": value1,
                            "value2": value2
                        })  
            else:
                if subject not in differences:
                    differences[subject] = {}
                if name not in differences[subject]:
                    differences[subject][name] = []
                differences[subject][name].append({
                    "field": "không tồn tại",
                    "value1": None,
                    "value2": None
                })
                
    return differences

def print_pretty_differences(differences):
    for subject, students in differences.items():
        print(f"- {subject} có {len(students)} học sinh có sự khác biệt")
        for student, fields in students.items():
            field_differences = "; ".join(
                [
                    f"{field['field']} ({field['value1']} != {field['value2']})"
                    if not (field['value1'] is None and field['value2'] is None)
                    else field['field']
                    for field in fields
                ]
            )
            print(f" + {student} {field_differences}")    
    
def export_differences_to_txt(differences, output_file):
    try:
        with open(output_file, "w") as file:
            if not differences:
                file.write("Không khác nhau.\n")
            else:
                for subject, students in differences.items():
                    file.write(f"- {subject} có {len(students)} học sinh có sự khác biệt\n")
                    for student, fields in students.items():
                        field_differences = "; ".join(
                            [
                                f"{field['field']} ({field['value1']} != {field['value2']})"
                                if not (field['value1'] is None and field['value2'] is None)
                                else field['field']
                                for field in fields
                            ]
                        )
                        file.write(f" + {student} {field_differences}\n")
                    
        print(f"Differences exported to {output_file}")
    except Exception as e:
        print(f"Error exporting differences to file: {e}")

    