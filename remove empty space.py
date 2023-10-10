import openpyxl
# 去除Excel表格中B列单元格开头的空格

# Load the workbook from D:\工作簿1.xlsx
path = "D:\\工作簿1.xlsx"
wb = openpyxl.load_workbook(path)
sheet = wb.active

# Iterate through all cells in column B
for row in sheet.iter_rows(min_col=2, max_col=2):
    for cell in row:
        if cell.value and isinstance(cell.value, str) and cell.value.startswith(" "):
            cell.value = cell.value[1:]

# Save the modified workbook to D:\工作簿1_modified.xlsx
output_path = "D:\\工作簿1_modified.xlsx"
wb.save(output_path)
