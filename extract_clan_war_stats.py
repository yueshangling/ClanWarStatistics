# -*- coding: utf-8 -*-
import os
import json
import pandas as pd

# 定义要读取的 JSON 文件所在目录
directory = 'ClanCompetitionStatistics/2025/'  # 替换为你的实际路径

# 创建一个空的 DataFrame 来存储统计结果
data = {
    '名称': [],
    '1星': [],
    '2星': [],
    '3星': [],
    '总进攻次数': [],
    '总获得星星': []
}

df = pd.DataFrame(data)

# 定义一个递归函数来查找并处理文件
def process_files(dir_path):
    for filename in os.listdir(dir_path):
        filepath = os.path.join(dir_path, filename)
        if os.path.isdir(filepath):
            # 如果是目录，则递归处理
            process_files(filepath)
        elif '统计结果' in filename and filename.endswith('.json'):
            # 如果是 JSON 文件，则处理
            print(f"正在处理 {filepath}...")
            with open(filepath, 'r', encoding='utf-8') as file:
                json_data = json.load(file)
                if isinstance(json_data, list):
                    # 如果 json_data 是一个列表，则遍历列表中的每个字典
                    for item in json_data:
                        process_json_item(item)
                elif isinstance(json_data, dict):
                    # 如果 json_data 是一个字典，则直接处理
                    process_json_item(json_data)

def process_json_item(json_item):
    # 检查 JSON 对象中是否包含所有需要的键
    required_keys = ['名称', '1星', '2星', '3星', '总进攻次数', '总获得星星']
    if not all(key in json_item for key in required_keys):
        # print(f"JSON 文件缺少必需的键: {json_item}")
        return

    # 检查 DataFrame 中是否已经存在相同名称的记录
    if json_item['名称'] in df['名称'].values:
        # 找到对应索引
        index = df[df['名称'] == json_item['名称']].index[0]
        # 合并数据
        df.at[index, '1星'] += json_item['1星']
        df.at[index, '2星'] += json_item['2星']
        df.at[index, '3星'] += json_item['3星']
        df.at[index, '总进攻次数'] += json_item['总进攻次数']
        df.at[index, '总获得星星'] += json_item.get('总获得星星','0')
    else:
        # 如果不存在，直接添加新记录
        df.loc[len(df)] = {
            '名称': json_item['名称'],
            '1星': json_item['1星'],
            '2星': json_item['2星'],
            '3星': json_item['3星'],
            '总进攻次数': json_item['总进攻次数'],
            '总获得星星': json_item['总获得星星']
        }

# 调用递归函数处理文件
process_files(directory)

# 将 DataFrame 写入 Excel 文件
output_file = '统计结果.xlsx'
df.to_excel(output_file, index=False, engine='openpyxl')

print(f"统计结果已保存到 {output_file}")
