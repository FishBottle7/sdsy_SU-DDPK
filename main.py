import openpyxl
import tqdm
import os

# Load the workbook
def load_workbook():
    path = input("输入表格路径：")
    wb = openpyxl.load_workbook(path)
    return wb, path

#拆分合并单元格
def unmerge_cells(wb):
    merged_cells_ranges = list(wb.active.merged_cells.ranges)  # Create a copy

    # Loop through all merged cells
    for merged_cells_range in tqdm.tqdm(merged_cells_ranges, desc="拆分合并单元格"):
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
    wb.save(new_file_path)
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

def format_grade_and_class(wb):
    sheet = wb.active
    grade_mapping = {"高一": "1", "高二": "2", "高三": "3"}  # Add a mapping from grade names to numbers

    # Iterate over rows
    for row in sheet.iter_rows(min_row=2):  # Skip header row
        grade = grade_mapping.get(row[2].value, row[2].value)  # Use the mapping to convert grade names to numbers
        class_num = row[3].value
        formatted_value = f"{grade},{class_num}"
        row[7].value = formatted_value  # Write to the 8th column (0-indexed)

    return wb

if __name__ == "__main__":
    wb, path = load_workbook()
    wb = unmerge_cells(wb)
    wb = keep_max_rows(wb)
    wb = remove_zeros(wb)
    wb = calculate_scores(wb)
    wb = format_grade_and_class(wb)
    save_workbook(wb, path)