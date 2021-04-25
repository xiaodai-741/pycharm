# !/usr/bin/env python
import sys
from xlrd import open_workbook  # xlrd用于读取xld
import xlwt  # 用于写入xls

workbook = open_workbook(r'1_1_原有老客户合同续签申请表-成都壹万佳茶文化传播有限公司.xls')  # 打开xls文件
sheet_name = workbook.sheet_names()  # 打印所有sheet名称，是个列表
sheet = workbook.sheet_by_index(0)  # 根据sheet索引读取sheet中的所有内容
print(sheet.name, sheet.nrows, sheet.ncols)  # sheet的名称、行数、列数
content = sheet.cell(1, 0)  # 第六列内容
print(content.value)
