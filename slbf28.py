#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
@Author:Alan
@Create 2023/8/8 21:06
@Version 1.0
"""
import os

import openpyxl

# 指定文件夹路径

folder_path = './testExcelFile'

# 创建一个新的工作簿用于存储提取的数据
new_workbook = openpyxl.Workbook()
new_sheet = new_workbook.active
new_sheet.append(["id", "type_classe", "tablename"])

print(os.listdir(folder_path))

# 遍历指定文件夹下的所有xlsx文件
for filename in os.listdir(folder_path):
    if filename.endswith('.xlsx'):
        source_filepath = os.path.join(folder_path, filename)

        # 打开源工作簿
        source_workbook = openpyxl.load_workbook(source_filepath)
        source_sheet = source_workbook.active

        # 获取工作簿最大行数
        max_row = source_sheet.max_row

        # 读取第一行，确定id和type_classe所在的列
        header_row = list(source_sheet.iter_rows(min_row=1, max_row=1, values_only=True))[0]
        id_col = 1  # 索引从 1 开始
        type_classe_col = None
        for col_index, header in enumerate(header_row):
            if header == "type_classe":
                type_classe_col = col_index + 1  # 列索引从1开始

        # 遍历处第一行外每一行的第一个单元格确定背景填充色
        for i in range(2, max_row + 1):
            # 获取表格 填充色 颜色
            color = source_sheet.cell(i, 1).fill.fgColor.rgb
            # 获取黄色行的id和type_classe
            if color == "FFFFFF00":
                id_value = source_sheet.cell(i, id_col).value
                type_classe_value = source_sheet.cell(i, type_classe_col).value
                new_sheet.append([id_value, type_classe_value])

        source_workbook.close()

# 保存新的工作簿到新文件
new_workbook.save('filtered_data.xlsx')
