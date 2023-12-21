# -*- coding: utf-8 -*-
"""
Created on Thu Nov  9 13:44:47 2023

@author: Andrew Liao
"""

import os
import pandas as pd

# 指定包含XLSX文件的文件夹路径
folder_path = r"E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Water temperature & flow rate and heating quantity\Raw data\2022"

# 创建一个ExcelWriter对象，用于写入XLSX文件
output_file = r"E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Water temperature & flow rate and heating quantity\Industrial and education building\2022.xlsx"
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# 定义要筛选的条件和对应的sheet名称
conditions = {
    '60404': 'Supply water temperature',
    '60405': 'Return water temperature',
    '60406': 'Water flow',
    '60425': 'Total heating quantity',
    '60409': 'Cooling quantity',
    '60410': 'Heating quantity'
}

# 遍历筛选条件，并创建对应的sheet
for condition, sheet_name in conditions.items():
    df_concat = pd.DataFrame()  # 创建一个空的DataFrame来存储筛选数据

    # 遍历文件夹中的所有XLSX文件
    for filename in os.listdir(folder_path):
        if filename.endswith(".xlsx"):
            file_path = os.path.join(folder_path, filename)

            # 读取XLSX文件
            try:
                df = pd.read_excel(file_path)

                # 确保列数大于等于1
                if df.shape[1] >= 1:
                    selected_data = df[df.iloc[:, 0].astype(str) == condition]  # 根据条件筛选数据
                    if not selected_data.empty:
                        # 在selected_data DataFrame中添加一个新的列以标识来源文件
                        selected_data['SourceFile'] = filename
                        # 使用前一个数据替代错误数据
                        selected_data.iloc[:, 1:] = selected_data.iloc[:, 1:].fillna(method='ffill')
                        # 将当前筛选数据追加到df_concat DataFrame中
                        df_concat = df_concat.append(selected_data, ignore_index=True)
                else:
                    print("文件 {} 中的列数小于1".format(file_path))

                print("文件 {} 的筛选数据已添加到DataFrame。".format(file_path))

            except pd.errors.ParserError:
                print("在文件 {} 中遇到错误，请检查文件格式。".format(file_path))

    # 将当前条件的筛选数据保存到指定的sheet
    df_concat.to_excel(writer, sheet_name=sheet_name, index=False)
    print("条件 {} 的筛选数据已保存到sheet {}。".format(condition, sheet_name))

# 保存并关闭ExcelWriter对象
writer.save()
writer.close()

print("所有文件的筛选数据已成功保存到不同sheet的文件:", output_file)