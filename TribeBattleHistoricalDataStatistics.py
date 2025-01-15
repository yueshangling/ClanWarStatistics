import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
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
                    first_attack_detail = entry.get("第一次攻击详情", "未使用")
                    if first_attack_detail != "未使用":
                        stars = count_stars(first_attack_detail)
                        stats["star_counts"][stars] += 1
                    else:
                        stats["unused_attacks"] += 1

                    # 统计第二次攻击
                    stats["total_attacks"] += 1
                    stats["second_attack_total"] += 1
                    second_attack_detail = entry.get("第二次攻击详情", "未使用")
                    if second_attack_detail != "未使用":
                        stars = count_stars(second_attack_detail)
                        stats["star_counts"][stars] += 1
                    else:
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

def export_to_excel_with_styles(results, output_path):
    """将结果导出到 Excel 并设置样式"""
    wb = Workbook()
    ws = wb.active
    ws.title = "统计结果"

    # 定义表头
    headers_player = [
        "名称", "1星", "1星占比", "2星", "2星占比", "3星", "3星占比", 
        "黑三", "黑三占比", "总进攻次数", "未使用进攻次数", 
        "未使用进攻次数占比"
    ]
    ws.append(headers_player)

    # 设置表头样式
    header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
    header_font = Font(bold=True, size=14, color="000000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_border = Border(
        top=Side(style="thin"),
        left=Side(style="thin"),
        bottom=Side(style="thin"),
        right=Side(style="thin")
    )
    for col_num, col_header in enumerate(headers_player, start=1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_alignment
        cell.border = header_border

    # 填充数据
    for result in results:
        row = [
            result["名称"],
            result["1星"],
            result["1星占比"],  # 确保是浮点数
            result["2星"],
            result["2星占比"],  # 确保是浮点数
            result["3星"],
            result["3星占比"],  # 确保是浮点数
            result["黑三"],
            result["黑三占比"],  # 确保是浮点数
            result["总进攻次数"],
            result["未使用进攻次数"],
            result["未使用进攻次数占比"],  # 确保是浮点数
        ]
        ws.append(row)

    # 设置百分比列格式
    percent_columns = [3, 5, 7, 9, 12]
    for col in percent_columns:
        for cell in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col, max_col=col):
            cell[0].number_format = "0.00%"  # 设置为百分比格式，且存储为数值

    # 标红未使用进攻次数等于总进攻次数的行
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=10, max_col=11):
        total_attacks_cell, unused_attacks_cell = row
        if total_attacks_cell.value == unused_attacks_cell.value:
            for cell in row:
                cell.fill = red_fill

    wb.save(output_path)
    print(f"统计结果已保存到: {output_path}")


# 调用函数
folder_path = "TribeBattleHistoricalData"  # 替换为你的 JSON 文件夹路径
output_path = "部落战统计.xlsx"  # 输出文件路径
results = analyze_folder_by_name(folder_path)
export_to_excel_with_styles(results, output_path)
