import os
import pandas as pd
from openpyxl import Workbook, load_workbook

# 初始化一个新的Excel工作簿
new_wb = Workbook()
new_ws = new_wb.active
new_ws.title = "Aggregated Data"

# 为新工作簿设置列标题（如果需要）
new_ws['A1'] = 'File Name'
new_ws['B1'] = 'A1'
new_ws['C1'] = 'A2'
new_ws['D1'] = 'A3'

# 初始行号
row_number = 2

# 遍历包含Excel文件的目录
for filename in os.listdir('.'):
    if filename.endswith('.xlsx') and filename != 'new_file.xlsx':  # 避免读取新文件（如果它已经存在）
        print(f"Reading {filename}")

        # 加载单个Excel文件
        wb = load_workbook(filename)
        ws = wb.active  # 默认读取第一个工作表

        # 从A1, A2, A3读取数据
        a1_data = ws['A1'].value
        a2_data = ws['A2'].value
        a3_data = ws['A3'].value

        # 将数据写入新的Excel文件
        new_ws[f'A{row_number}'] = filename
        new_ws[f'B{row_number}'] = a1_data
        new_ws[f'C{row_number}'] = a2_data
        new_ws[f'D{row_number}'] = a3_data

        # 更新行号
        row_number += 1

# 保存新的Excel文件
new_wb.save('new_file.xlsx')
print("Aggregated data saved to 'new_file.xlsx'")
