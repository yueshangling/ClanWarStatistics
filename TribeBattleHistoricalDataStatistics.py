import os
import json
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from collections import defaultdict

headers_player = [
    "名称",
    "1星",
    "1星占比",
    "2星",
    "2星占比",
    "3星占比",
    "黑三",
    "黑三占比",
    "总进攻次数",
    "未使用进攻次数",
    "未使用进攻次数占比",
    "第一次攻击占比",
    "第二次攻击占比",
]
# 新增功能：收集所有两次进攻未使用的记录
def collect_unused_attacks(folder_path):
    """收集所有两次攻击都未使用的记录"""
    dir_unused = defaultdict(list)  # 按目录分组的记录
    total_unused = []               # 总记录
    
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file.endswith('.json'):
                file_path = os.path.join(root, file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    try:
                        data = json.load(f)
                        for entry in data:
                            # 检查两次攻击是否都未使用
                            first = entry.get("第一次攻击详情", "未使用")
                            second = entry.get("第二次攻击详情", "未使用")
                            if first == "未使用" and second == "未使用":
                                # 获取相对路径用于分组
                                relative_path = os.path.relpath(root, folder_path)
                                record = {
                                    "名称": entry.get("名称", "未知"),
                                    "第一次攻击详情": first,
                                    "第二次攻击详情": second,
                                    "来源文件": os.path.basename(file_path),
                                    "来源目录": relative_path
                                }
                                dir_unused[relative_path].append(record)
                                total_unused.append(record)
                    except json.JSONDecodeError:
                        print(f"文件解析失败: {file_path}")
    return dir_unused, total_unused
def export_unused_records(dir_unused, total_unused, base_folder):
    """导出未使用记录到对应目录和总文件"""
    # 导出各目录的记录
    for relative_path, records in dir_unused.items():
        if records:
            # 创建目录路径
            dir_path = os.path.join(base_folder, relative_path)
            os.makedirs(dir_path, exist_ok=True)
            
            # 生成 Excel 文件
            output_path = os.path.join(dir_path, "两次未使用记录.xlsx")
            df = pd.DataFrame(records)
            df.to_excel(output_path, index=False)
            print(f"已生成目录记录文件: {output_path}")
    
    # 导出总记录
    if total_unused:
        output_path = os.path.join(base_folder, "总未使用记录.xlsx")
        df = pd.DataFrame(total_unused)
        df.to_excel(output_path, index=False)
        print(f"已生成总记录文件: {output_path}")
    else:
        print("没有未使用的进攻记录")
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
def mark_rows_red(ws):
    # 定义红色填充样式
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    # 遍历从第二行开始的每一行，从第 10 列到第 11 列
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=10, max_col=11):
        total_attacks_cell, unused_attacks_cell = row
        if total_attacks_cell.value == unused_attacks_cell.value:
            # 获取整行的单元格范围
            full_row = ws[row[0].row]
            for cell in full_row:
                cell.fill = red_fill
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
            float(result["1星占比"].replace("%", "")) / 100,  # 将百分比字符串转换为浮点数
            result["2星"],
            float(result["2星占比"].replace("%", "")) / 100,  # 将百分比字符串转换为浮点数
            result["3星"],
            float(result["3星占比"].replace("%", "")) / 100,  # 将百分比字符串转换为浮点数
            result["黑三"],
            float(result["黑三占比"].replace("%", "")) / 100,  # 将百分比字符串转换为浮点数
            result["总进攻次数"],
            result["未使用进攻次数"],
            float(result["未使用进攻次数占比"].replace("%", "")) / 100,  # 将百分比字符串转换为浮点数
        ]
        ws.append(row)

    # 设置百分比列格式
    percent_columns = [3, 5, 7, 9, 12]
    for col in percent_columns:
        for cell in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=col, max_col=col):
            cell[0].number_format = "0.00%"  # 设置为百分比格式，且存储为数值
    # 标红未使用进攻次数等于总进攻次数的行
    mark_rows_red(ws)
    adjust_column_width(ws)
    wb.save(output_path)
    print(f"统计结果已保存到: {output_path}")
def adjust_column_width(ws):
    max_lengths = {}
    for i, header in enumerate(headers_player):
        col_letter = chr(65 + i)
        max_lengths[col_letter] = {'value': len(header), 'type': 'str'}
    for row in ws.iter_rows():
        for cell in row:
            col_letter = cell.column_letter
            try:
                if isinstance(cell.value, float):  # 如果是数字
                    max_lengths[col_letter]['type'] = 'int'
                max_lengths[col_letter]['value'] = max(max_lengths[col_letter]['value'], len(str(cell.value)))
            except:
                max_lengths[col_letter] = {'value': len(str(cell.value)), 'type': 'str'}
    for col, length in max_lengths.items():
        if length['type'] == 'int':
            ws.column_dimensions[col].width = length['value'] + 4  # 增加数字列的宽度
        else:
            ws.column_dimensions[col].width = length['value'] * 2.5 + 2
# 调用函数
folder_path = "TribeBattleHistoricalData"  # 替换为你的 JSON 文件夹路径
output_path = "部落战统计.xlsx"  # 输出文件路径
results = analyze_folder_by_name(folder_path)
export_to_excel_with_styles(results, output_path)
# 新增未使用记录统计
dir_unused, total_unused = collect_unused_attacks(folder_path)
export_unused_records(dir_unused, total_unused, folder_path)
