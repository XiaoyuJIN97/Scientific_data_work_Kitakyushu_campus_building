# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 15:15:24 2023

@author: Andrew Liao
"""

import numpy as np
import pandas as pd
import datetime

# 读取Excel文件# 读取Excel文件
file_path = r"E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Heating load & Cooling load & Hot water load\2022.xlsx"  # 请将文件路径更新为实际路径

sheet_name1 = 0
df = pd.read_excel(file_path, sheet_name= sheet_name1)
sheet_name2 = 1
df2 = pd.read_excel(file_path, sheet_name= sheet_name2)
sheet_name3 = 2
df3 = pd.read_excel(file_path, sheet_name = sheet_name3)

# 提取数据
data = df.iloc[0:366, 6:30]
data2 = df2.iloc[0:366, 6:30]
data3 = df3.iloc[0:366, 6:30]

# 将空缺值替换为中位数
data = data.fillna(data.median())
data2 = data2.fillna(data2.median())
data3 = data3.fillna(data3.median())

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

# 将数据展平为一列
processed_data = data.values.flatten()
processed_data2 = data2.values.flatten()
processed_data3 = data3.values.flatten()

# 创建包含日期和数据的DataFrame
processed_df = pd.DataFrame({'Date': np.repeat(dates, len(times)), 'Cooling load (kWh)': processed_data, 'Heating load (kWh)': processed_data2, 'Hot water load (kWh)': processed_data3})

# 保存处理后的数据为Excel文件
processed_df.to_excel(r'E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\Heating load & Cooling load & Hot water load\Time series_2022.xlsx', index=False)

# 输出结果
print("数据已处理并保存为 processed_data.xlsx 文件。")