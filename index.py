import re
from openpyxl import load_workbook



# 加载 Excel 文件
workbook = load_workbook('file.xlsx')
sheet = workbook.active
项目号 = sheet.cell(row=3, column=4).value
项目名称 = sheet.cell(row=2, column=4).value
业主名称 = sheet.cell(row=2, column=11).value
print(项目号)
print(项目名称)

# 获取从第四行开始的所有内容
data = []
for row in sheet.iter_rows(min_row=5, values_only=True):
    if (row[0]):
        data.append([项目名称, row[3], 项目号, "", re.findall(r'\d+', 项目名称)[0], 业主名称, row[7],row[6],row[8], None, None, None, row[12], None,None,None,row[15],None,None,row[16],None,None, row[17],None,None,row[18],None,None,row[19],None,None, row[20]])


# 加载 Excel 文件
workbook = load_workbook('example.xlsx')

# 选择活动工作表
sheet = workbook.active

for new_row in data:

    # 在第三行添加一行数据，注意openpyxl的行和列索引从1开始
    sheet.insert_rows(13)
    for col_num, value in enumerate(new_row, start=1):
        sheet.cell(row=13, column=col_num, value=value)

# 保存文件
workbook.save('example.xlsx')