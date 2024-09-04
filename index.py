import re
import os
import json
import datetime
import requests
from openpyxl import Workbook
import numpy as np
from openpyxl import load_workbook
from flask import Flask, send_file, request

def loadExcel(file):
    # 加载 Excel 文件
    workbook = load_workbook(file)
    sheet = workbook.active
    项目号 = sheet.cell(row=3, column=4).value
    项目名称 = sheet.cell(row=2, column=4).value
    业主名称 = sheet.cell(row=2, column=11).value
    print(str(项目号) + ':' + 项目名称)

    # 获取从第四行开始的所有内容
    data = []
    # 定义替换函数
    def remove_zero_time(time):
        if (isinstance(time, datetime.datetime)):
            return time.strftime('%Y/%m/%d')
        return time
    for row in sheet.iter_rows(min_row=5, values_only=True):
        if (row[0]):
            lestTemp = [项目名称, row[3], 项目号, "", re.findall(r'\d+', 项目名称)[0], 业主名称, row[7],row[6],row[8], None, None, None, row[12], None,None,None,row[15],None,None,row[16],None,None, row[17],None,None,row[18],None,None,row[19],None,None, row[20]]
            lestTemp = list(map(remove_zero_time, lestTemp))
            data.append(lestTemp)
    
    # print(data)
    return data
    # # 加载 Excel 文件
    # workbook = load_workbook('main.xlsx')

    # # 选择活动工作表
    # sheet = workbook.active

    # for new_row in data:
    #     # 在第三行添加一行数据，注意openpyxl的行和列索引从1开始
    #     sheet.insert_rows(3)
    #     for col_num, value in enumerate(new_row, start=1):
    #         sheet.cell(row=3, column=col_num, value=value)

    # # 保存文件
    # workbook.save('main.xlsx')

# loadExcel()
# 指定目录路径
directory_path = './files'

# 获取目录下所有xlsx文件的名称
xlsx_files = [f for f in os.listdir(directory_path) if "立项审批流程" in f]

# outPutData = [["项目名称", "物料编码", "项目编号", "", "项目ID", "业主名称", "设备名称", "图号或规格型号", "合同数量", "", "", "", "合同交货日期", "", "", "", "图纸下达", "", "", "预算下达", "", "", "采购合同完成", "", "", "制造完成", "", "", "成品检验完成", "", "", "发运完成"]]

outPutData = [["项目名称", "合同号", "项目号", "产品类型", "令号", "业主名称", "产品名称", "型号", "数量", "销售经理", "外委地点", "立项时间", "合同交付日", "排产时间", "是否交付", "所属进度", "图纸计划", "完成时间", "超期", "预算计划", "完成时间", "超期", "采购计划", "完成时间", "超期", "制造计划", "完成时间", "超期", "检验计划", "完成时间", "超期", "发运计划","完成时间","超期","预警","进度/风险提示","调整/确定交付日","总进度超期","超期原因","合同金额（万元）","罚款关注","罚款","合同付款方式","收款完成","待收款性质","累计已收款","已收款比例","未收款","已支付","可支付余额","所属板块","国内/海外"]]
for item in xlsx_files:
    print('./files/' + item)
    outPutData.extend(loadExcel('./files/' + item))
print(outPutData)
# 保存为CSV文件
file_path = 'file.xlsx'

# 创建一个新的Workbook对象
wb = Workbook()
ws = wb.active  # 获取活动的工作表

# 将数组中的每一行写入工作表
for row in outPutData:
    ws.append(row)

# 保存为.xlsx文件
wb.save(file_path)

print(f"数组已保存为xlsx文件: {file_path}")