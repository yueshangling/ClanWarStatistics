# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import json
with open('data.json', 'r', encoding='utf-8') as file:
    game_data = json.load(file)

# 创建工作簿
workbook = Workbook()

# 处理我方队员进攻数据的工作表
worksheet_player = workbook.active
worksheet_player.title = "我方队员进攻"

# 定义表头数据
headers_player = ["序号", "名称", "职位", "部落等级", "第一次攻击", "第一次攻击详情", "第二次攻击", "第二次攻击详情",
                  "获得的星", "评价"]
worksheet_player.append(headers_player)

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
# 应用表头样式到每一个表头单元格
for col_num, col_header in enumerate(headers_player, start=1):
    cell = worksheet_player.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = header_alignment
    cell.border = header_border

# 填充我方队员进攻数据，并设置数据单元格样式
for index, data in enumerate(game_data, start=1):
    row_data = [index, data["名称"], data["职位"], data.get("分数", "已退出"), data["第一次攻击"],
                data["第一次攻击详情"], data["第二次攻击"], data["第二次攻击详情"], data["获得的星"],]
    worksheet_player.append(row_data)
    for col_num in range(1, len(row_data) + 1):
        cell = worksheet_player.cell(row=worksheet_player.max_row, column=col_num)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        # 判断是否两个字段值都是"未使用"，若是则整行标红
        if data["第一次攻击详情"] == "未使用" and data["第二次攻击详情"] == "未使用":
            red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            for col in range(1, len(headers_player) + 1):
                target_cell = worksheet_player.cell(row=worksheet_player.max_row, column=col)
                target_cell.fill = red_fill
        cell.border = Border(
            top=Side(style="thin"),
            left=Side(style="thin"),
            bottom=Side(style="thin"),
            right=Side(style="thin")

        )

# 保存工作簿
workbook.save("game_data.xlsx")
print("数据已成功写入game_data.xlsx文件，且设置了相应的样式。")
