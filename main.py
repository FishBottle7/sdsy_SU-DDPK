import openpyxl
import os
import atexit

# Load the workbook
def load_workbook():
    path = input("输入表格路径：")
    path = path.strip('\'"')
    while True:
        try:
            wb = openpyxl.load_workbook(path)
            break
        except PermissionError:
            input("请关闭文件后重试")
    return wb, path

#拆分合并单元格
def unmerge_cells(wb):
    merged_cells_ranges = list(wb.active.merged_cells.ranges)  # Create a copy

    # Loop through all merged cells
    for merged_cells_range in merged_cells_ranges:
        # Get the top-left cell value
        top_left_cell_value = merged_cells_range.start_cell.value

        # Unmerge cells
        wb.active.unmerge_cells(str(merged_cells_range))

        # Fill unmerged cells with the top-left cell value
        for row in wb.active[merged_cells_range.coord]:
            for cell in row:
                cell.value = top_left_cell_value

    return wb

def save_workbook(wb, path):
    file_path, extension = os.path.splitext(path)

    # Add 'processed' to the file path
    new_file_path = f"{file_path}_processed{extension}"

    # Save the workbook with the new file path
    try:
        wb.save(new_file_path)
    except PermissionError:
        input("请关闭文件后重试")
        return save_workbook(wb, path)
    print(f"Workbook saved at {new_file_path}")

def keep_max_rows(wb):
    sheet = wb.active
    max_rows = {}
    max_seq = {}  # Initialize max sequence number for each name
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
        name = row[1]
        try:
            seq = int(row[0])  # Convert '序号' to int
        except Exception:
            continue  # Skip rows where '序号' cannot be converted to int
        if name not in max_seq or seq > max_seq[name]:  # If sequence number is greater than max sequence number for this name
            max_rows[name] = [row]  # Reset rows for this name
            max_seq[name] = seq  # Update max sequence number for this name
        elif seq == max_seq[name]:  # If sequence number is equal to max sequence number for this name
            max_rows[name].append(row)  # Add row to list of rows for this name

    # Create a new workbook with only the max rows
    new_wb = openpyxl.Workbook()
    new_sheet = new_wb.active
    new_sheet.append(list(next(sheet.iter_rows(min_row=1, max_row=1, values_only=True))))  # Copy header row
    for rows in max_rows.values():
        for row in rows:
            new_sheet.append(row)

    return new_wb

def remove_zeros(wb):
    sheet = wb.active
    rows_to_delete = []  # List to keep track of rows to delete

    # Iterate over rows
    for i, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):  # Skip header row
        if float(row[4]) == 0:  # If value in the fifth column is 0
            rows_to_delete.append(i)

    # Delete rows in reverse order to avoid shifting indices
    for i in reversed(rows_to_delete):
        sheet.delete_rows(i)

    return wb

def absolute_values(wb):
    sheet = wb.active

    for row in sheet.iter_rows(min_row=2):  # Skip header row
        try:
            value = float(row[4].value)
            row[4].value = abs(value)
        except ValueError:
            continue

    return wb

def calculate_scores(wb):
    sheet = wb.active

    # Set headers for columns 13 and 14
    sheet.cell(row=1, column=13, value="班级")
    sheet.cell(row=1, column=14, value="总分")

    # Initialize class names and scores
    class_names = [f"高一（{i}）班" for i in range(1, 14)] + [f"高二（{i}）班" for i in range(1, 14)] + [f"高三（{i}）班" for i in range(1, 14)]
    class_scores = {class_name: 0 for class_name in class_names}

    # Calculate scores for each class
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
        grade = row[2]
        class_num = row[3]
        score = float(row[4])
        class_name = f"{grade}（{class_num}）班"
        if class_name in class_scores:
            class_scores[class_name] -= score  # Subtract score because we want the opposite of the sum

    # Write class names and scores to columns 13 and 14
    for i, (class_name, score) in enumerate(class_scores.items(), start=2):  # Skip header row
        sheet.cell(row=i, column=13, value=class_name)
        sheet.cell(row=i, column=14, value=score)

    return wb

def calculate_scores_with_classify(wb):
    sheet = wb.active

    # Set headers for columns 13, 14, 15, 16, 17
    sheet.cell(row=1, column=13, value="班级")
    sheet.cell(row=1, column=14, value="升旗")
    sheet.cell(row=1, column=15, value="两操")
    sheet.cell(row=1, column=16, value="日常")
    sheet.cell(row=1, column=17, value="周五检查")

    # Initialize class names and scores
    class_names = [f"高一（{i}）班" for i in range(1, 14)] + ["1"] + [f"高二（{i}）班" for i in range(1, 14)] + ["2"] + [f"高三（{i}）班"for i in range(1, 14)]
    class_category_scores = {class_name: {"升旗": None, "两操": None, "日常": None, "周五检查": None} if class_name == "1" or class_name == "2"
                                        else {"升旗": 0, "两操": 0, "日常": 0, "周五检查": 0} for class_name in class_names}
    # Calculate scores for each class
    for row in sheet.iter_rows(min_row=2, values_only=True):  # Skip header row
        grade = row[2]
        class_num = row[3].replace("班", "")  # Remove '班' from class number
        score = float(row[4])
        category = row[6]
        class_name = f"{grade}（{class_num}）班"
        if class_name in class_category_scores:
            if category in class_category_scores[class_name]:
                class_category_scores[class_name][category] -= score  # Subtract score because we want the opposite of the sum
            else:
                print(f"line {row} Unexpected category: {category}")
                continue
        else:
            print(f"line {row} Unexpected class name: {class_name}")
            continue

    # Write class names to column 13
    for i, class_name in enumerate(class_names, start=2):  # Skip header row
        sheet.cell(row=i, column=13, value=class_name)


    # Write class category scores to columns 14, 15, 16, 17
    for i, (class_name, category_scores) in enumerate(class_category_scores.items(), start=2):  # Skip header row
        sheet.cell(row=i, column=14, value=category_scores["升旗"])
        sheet.cell(row=i, column=15, value=category_scores["两操"])
        sheet.cell(row=i, column=16, value=category_scores["日常"])
        sheet.cell(row=i, column=17, value=category_scores["周五检查"])

    return wb

def format_grade_and_class(wb):
    sheet = wb.active
    grade_mapping = {"高一": "1", "高二": "2", "高三": "3"}  # Add a mapping from grade names to numbers

    # Iterate over rows
    for row in sheet.iter_rows(min_row=2):  # Skip header row
        grade = grade_mapping.get(row[2].value, row[2].value)  # Use the mapping to convert grade names to numbers
        #class_num = row[3].value.replace("班", "")  # Remove '班' from class number
        class_num = row[3].value.replace("班", "") if row[3].value else None
        if class_num is None: continue
        formatted_value = f"{grade},{class_num}"
        row[7].value = formatted_value  # Write to the 8th column (0-indexed)

    return wb

def exit_handler():
    input("按任意键退出")

atexit.register(exit_handler)

if __name__ == "__main__":
    wb, path = load_workbook()
    wb = unmerge_cells(wb)
    wb = keep_max_rows(wb)
    wb = remove_zeros(wb)
    wb = absolute_values(wb)
    wb = calculate_scores_with_classify(wb)
    wb = format_grade_and_class(wb)
    save_workbook(wb, path)