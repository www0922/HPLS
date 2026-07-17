"""
统一格式: 所有单元格字体设为 Times New Roman
"""
import openpyxl
from pathlib import Path
from openpyxl.styles import Font

BASE = Path(__file__).resolve().parent.parent / 'excel_data'
DST = str(BASE / '20260711文库pooling表T7+PE150-zss.xlsx')
FONT = Font(name='Times New Roman', size=10)

wb = openpyxl.load_workbook(DST)

for sn in wb.sheetnames:
    ws = wb[sn]
    for row in ws.iter_rows(min_row=2):  # 跳过表头行
        for cell in row:
            cell.font = FONT
    print(f'{sn}: 字体已设置')

wb.save(DST)
print(f'\n✅ 完成 → {DST}')
