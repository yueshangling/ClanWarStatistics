import os
import json
import pandas as pd
from collections import defaultdict

def count_stars(details):
    """统计星星数"""
    if details == "未使用":
        return 0
    return details.count('★')

def analyze_folder_by_name(folder_path):
    stats_by_name = defaultdict(lambda: {
        "total_attacks": 0,
        "unused_attacks": 0,
        "first_attack_total": 0,
        "second_attack_total": 0,
        "star_counts": defaultdict(int),
    })

    def process_file(file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            try:
                data = json.load(file)
                for entry in data:
                    name = entry.get("名称", "未知")
                    stats = stats_by_name[name]

                    # 统计第一次攻击
                    stats["total_attacks"] += 1
                    stats["first_attack_total"] += 1
                    stars = count_stars(entry.get("第一次攻击详情", "未使用"))
                    stats["star_counts"][stars] += 1
                    if stars == 0:
                        stats["unused_attacks"] += 1

                    # 统计第二次攻击
                    stats["total_attacks"] += 1
                    stats["second_attack_total"] += 1
                    stars = count_stars(entry.get("第二次攻击详情", "未使用"))
                    stats["star_counts"][stars] += 1
                    if stars == 0:
                        stats["unused_attacks"] += 1
            except json.JSONDecodeError:
                print(f"文件解析失败: {file_path}")

    def recursive_read_folder(folder):
        for root, _, files in os.walk(folder):
            for file in files:
                if file.endswith('.json'):
                    process_file(os.path.join(root, file))

    recursive_read_folder(folder_path)

    # 计算统计结果
    results = []
    for name, stats in stats_by_name.items():
        total_attacks = stats["total_attacks"]
        unused_attacks = stats["unused_attacks"]
        first_attack_total = stats["first_attack_total"]
        second_attack_total = stats["second_attack_total"]

        # 各星级统计
        stars_1 = stats["star_counts"][1]
        stars_2 = stats["star_counts"][2]
        stars_3 = stats["star_counts"][3]
        stars_0 = stats["star_counts"][0]

        results.append({
            "名称": name,
            "1星": stars_1,
            "1星占比": f"{stars_1 / total_attacks:.2%}" if total_attacks else "0.00%",
            "2星": stars_2,
            "2星占比": f"{stars_2 / total_attacks:.2%}" if total_attacks else "0.00%",
            "3星": stars_3,
            "3星占比": f"{stars_3 / total_attacks:.2%}" if total_attacks else "0.00%",
            "黑三": stars_0,
            "黑三占比": f"{stars_0 / total_attacks:.2%}" if total_attacks else "0.00%",
            "总进攻次数": total_attacks,
            "未使用进攻次数": unused_attacks,
            "未使用进攻次数占比": f"{unused_attacks / total_attacks:.2%}" if total_attacks else "0.00%",
            "第一次攻击占比": f"{first_attack_total / total_attacks:.2%}" if total_attacks else "0.00%",
            "第二次攻击占比": f"{second_attack_total / total_attacks:.2%}" if total_attacks else "0.00%",
        })

    return results

def export_to_excel_by_name(results, output_path):
    """将按名称统计的结果导出到 Excel"""
    df = pd.DataFrame(results)
    df.to_excel(output_path, index=False, sheet_name="统计结果", engine="openpyxl")
    print(f"统计结果已保存到: {output_path}")

# 调用函数
folder_path = "TribeBattleHistoricalData"  # 替换为你的 JSON 文件夹路径
output_path = "output_by_name.xlsx"  # 输出文件路径
results = analyze_folder_by_name(folder_path)
export_to_excel_by_name(results, output_path)
