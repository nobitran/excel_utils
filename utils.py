from collections import defaultdict

def get_col_index(col_letter):
    return ord(col_letter.upper()) - ord('A')

def get_subject_for_data2(mapping, subject):
    return mapping.get(subject, subject)

def compare_data(data1, data2, mapping):
    differences = defaultdict(lambda: defaultdict(list))
    for subject, records1 in data1.items():
        subject2 = get_subject_for_data2(mapping, subject)
        records2 = {r.get("name"): r for r in data2.get(subject2, [])}
        for record1 in records1:
            name = record1.get("name")
            record2 = records2.get(name)
            if record2:
                for key, value1 in record1.items():
                    value2 = record2.get(key)
                    try:
                        value1_f = float(value1)
                        value2_f = float(value2)
                        if value1_f != value2_f:
                            differences[subject][name].append({
                                "field": key, "value1": value1, "value2": value2
                            })
                    except (ValueError, TypeError):
                        if value1 != value2:
                            differences[subject][name].append({
                                "field": key, "value1": value1, "value2": value2
                            })
            else:
                differences[subject][name].append({
                    "field": "không tồn tại", "value1": None, "value2": None
                })
    return {k: dict(v) for k, v in differences.items()}

def print_pretty_differences(differences):
    for subject, students in differences.items():
        print(f"- {subject} có {len(students)} học sinh có sự khác biệt")
        for student, fields in students.items():
            field_differences = "; ".join(
                f"{f['field']} ({f['value1']} != {f['value2']})"
                if not (f['value1'] is None and f['value2'] is None)
                else f['field']
                for f in fields
            )
            print(f" + '{student}' {field_differences}")
        print("-" * 20)

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
                            f"{f['field']} ({f['value1']} != {f['value2']})"
                            if not (f['value1'] is None and f['value2'] is None)
                            else f['field']
                            for f in fields
                        )
                        file.write(f" + '{student}' {field_differences}\n")
        # print(f"Differences exported to {output_file}")
    except Exception as e:
        print(f"Error exporting differences to file: {e}")
