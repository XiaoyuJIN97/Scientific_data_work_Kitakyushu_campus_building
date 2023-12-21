# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 15:15:24 2023

@author: Andrew Liao
"""

import numpy as np
import pandas as pd
import datetime

# 读取Excel文件# 读取Excel文件
file_path = r"E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Water temperature & flow rate and heating quantity\Media building\2022.xlsx"  # 请将文件路径更新为实际路径

sheet_name1 = 0
df = pd.read_excel(file_path, sheet_name= sheet_name1)
sheet_name2 = 1
df2 = pd.read_excel(file_path, sheet_name= sheet_name2)
sheet_name3 = 2
df3 = pd.read_excel(file_path, sheet_name = sheet_name3)
sheet_name4 = 3
df4 = pd.read_excel(file_path, sheet_name= sheet_name4)
sheet_name5 = 4
df5 = pd.read_excel(file_path, sheet_name= sheet_name5)
sheet_name6 = 5
df6 = pd.read_excel(file_path, sheet_name = sheet_name6)

# 提取数据
data = df.iloc[0:366, 6:30]
data2 = df2.iloc[0:366, 6:30]
data3 = df3.iloc[0:366, 6:30]
data4 = df4.iloc[0:366, 6:30]
data5 = df5.iloc[0:366, 6:30]
data6 = df6.iloc[0:366, 6:30]

# 将空缺值替换为中位数
data = data.fillna(data.median())
data2 = data2.fillna(data2.median())
data3 = data3.fillna(data3.median())
data4 = data4.fillna(data4.median())
data5 = data5.fillna(data5.median())
data6 = data6.fillna(data6.median())

# 创建包含日期和时间的列名
dates = pd.date_range(start="2022-01-01", periods=304)
times = pd.date_range(start="00:00", end="23:00", freq="H").strftime("%H:%M")
columns = [date.strftime("%Y-%m-%d") + " " + time for date in dates for time in times]

# 调整列名数量以匹配数据列数
if len(data.columns) < len(columns):
    columns = columns[:len(data.columns)]
elif len(data.columns) > len(columns):
    data = data.iloc[:, :len(columns)]

if len(data2.columns) < len(columns):
    columns = columns[:len(data2.columns)]
elif len(data2.columns) > len(columns):
    data2 = data2.iloc[:, :len(columns)]

if len(data3.columns) < len(columns):
    columns = columns[:len(data3.columns)]
elif len(data3.columns) > len(columns):
    data3 = data3.iloc[:, :len(columns)]

if len(data4.columns) < len(columns):
    columns = columns[:len(data4.columns)]
elif len(data4.columns) > len(columns):
    data4 = data4.iloc[:, :len(columns)]

if len(data5.columns) < len(columns):
    columns = columns[:len(data5.columns)]
elif len(data5.columns) > len(columns):
    data5 = data5.iloc[:, :len(columns)]

if len(data6.columns) < len(columns):
    columns = columns[:len(data6.columns)]
elif len(data6.columns) > len(columns):
    data6 = data6.iloc[:, :len(columns)]

# 将数据展平为一列
processed_data = data.values.flatten()
processed_data2 = data2.values.flatten()
processed_data3 = data3.values.flatten()
processed_data4 = data4.values.flatten()
processed_data5 = data5.values.flatten()
processed_data6 = data6.values.flatten()

# 创建包含日期和数据的DataFrame
processed_df = pd.DataFrame({'Date': np.repeat(dates, len(times)), 'Supply water temperature (℃)': processed_data, 'Return water temperature (℃)': processed_data2, 'Water flow (m3/h)': processed_data3, 'Total heating quantity (GJ)': processed_data4, 'Cooling quantity (GJ)': processed_data5, 'Heating quantity (GJ)': processed_data6})

# 保存处理后的数据为Excel文件
processed_df.to_excel(r'E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Water temperature & flow rate and heating quantity\Media building\Time series_2022.xlsx', index=False)

# 输出结果
print("数据已处理并保存为 processed_data.xlsx 文件。")