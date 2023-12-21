# -*- coding: utf-8 -*-
"""
Created on Fri Nov 10 15:15:24 2023

@author: Andrew Liao
"""

import numpy as np
import pandas as pd
import datetime

# 读取Excel文件# 读取Excel文件
file_path = r"E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\AHU & N304\N304\N304_2009.xlsx"  # 请将文件路径更新为实际路径

sheet_name1 = 0
df = pd.read_excel(file_path, sheet_name= sheet_name1)

# 提取数据
data = df.iloc[0:367, 6:30]

# 将空缺值替换为1
data = data.fillna(data.median())

# 创建包含日期和时间的列名
dates = pd.date_range(start="2009-01-01", periods=365)
times = pd.date_range(start="00:00:00", end="23:00:00", freq="H").strftime("%H:%M")
columns = [date.strftime("%Y-%m-%d") + " " + time for date in dates for time in times]

# 调整列名数量以匹配数据列数
if len(data.columns) < len(columns):
    columns = columns[:len(data.columns)]
elif len(data.columns) > len(columns):
    data = data.iloc[:, :len(columns)]

# 将数据展平为一列
processed_data = data.values.flatten()


# 创建包含日期和数据的DataFrame
#processed_df = pd.DataFrame({'Date': np.repeat(dates, len(times)), 'N304': processed_data, 'Air supply temperature': processed_data2, 'Temperature below the corridor': processed_data3, 'Humidity below the corridor': processed_data4, 'Outdoor air intake temperature': processed_data5})

processed_df = pd.DataFrame({'Date': np.repeat(dates, len(times)), 'N304 temperature': processed_data})

# 保存处理后的数据为Excel文件
processed_df.to_excel(r'E:\Scientific_data_work_Kitakyushu_campus_building\Data Processing\AHU & N304\N304\N304_Time series_2009.xlsx', index=False)


# 输出结果
print("数据已处理并保存为 processed_data.xlsx 文件。")