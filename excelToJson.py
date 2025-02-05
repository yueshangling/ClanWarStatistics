# -*- coding: utf-8 -*-
import os
import pandas as pd
import json

# 定义文件夹路径
folder_path = 'TribeBattleHistoricalData/2025-1-31/'

# 递归遍历文件夹
for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.xlsx'):  # 只处理 Excel 文件
            file_path = os.path.join(root, file)
            print(f"正在处理文件: {file_path}")
            
            # 读取 Excel 文件
            df = pd.read_excel(file_path)

            # 将空值替换为空字符串
            df = df.astype(object)
            df.fillna('', inplace=True)
            
            # 构建 JSON 数据
            json_data = df.to_dict(orient='records')
            
            # 定义 JSON 文件路径
            json_file_path = os.path.join(root, f"{os.path.splitext(file)[0]}.json")
            
            # 将数据写入 JSON 文件
            with open(json_file_path, 'w', encoding='utf-8') as json_file:
                json.dump(json_data, json_file, ensure_ascii=False, indent=4)

            print(f"生成 JSON 文件: {json_file_path}")

print("所有 Excel 文件已转换为 JSON 文件")
