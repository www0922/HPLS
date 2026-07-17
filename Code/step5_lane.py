"""
步骤五：Lane编号、合并单元格、框线、文库环化填充
===============================================
1. B表 A/B/C 工作表:
   - A列写入 Lane编号 (A1,A2.../B1,B2.../C1,C2...)
   - 每组合并 A列 单元格
   - A列~K列 标框线
2. 文库环化 工作表:
   - 每组一行: A=Lane编号, E=汇总D, H=组第一行G, I=组第一行H
   - 公式: D=C*0.66*1, F=D, G=F/B, M=22*L/F
   - A列~K列 标框线
"""
import openpyxl
from pathlib import Path
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent.parent / 'excel_data'
DST = str(BASE / '20260711文库pooling表T7+PE150-zss.xlsx')
SHEETS = ['A', 'B', 'C']
CENTER = Alignment(horizontal='center', vertical='center')
ALL_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin'),
)

# 列索引
COL_A = 1
COL_D = 4
COL_G = 7
COL_H = 8
COL_K = 11


def read_groups(ws):
    """读取工作表, 按空白行分隔组, 每组含数据行+汇总行"""
    max_col = ws.max_column
    groups = []
    current = []

    for row in range(2, ws.max_row + 1):
        values = [ws.cell(row=row, column=c).value for c in range(1, max_col + 1)]
        if all(v is None for v in values):
            if current:
                groups.append(current)
                current = []
            continue
        current.append(row)

    if current:
        groups.append(current)
    return groups


def process_abc_sheet(ws, name):
    """处理 A/B/C 工作表: Lane编号、合并、框线"""
    groups = read_groups(ws)
    max_col = ws.max_column

    lane_info = []  # [(lane_name, first_g_val, first_h_val, summary_d_val), ...]

    for gi, rows in enumerate(groups):
        lane_name = f'{name}{gi + 1}'
        first_row = rows[0]
        last_row = rows[-1]

        # 取组内第一行的 G列、H列值
        first_g = ws.cell(row=first_row, column=COL_G).value
        first_h = ws.cell(row=first_row, column=COL_H).value
        # 取汇总行(最后一行) D列值
        summary_d = ws.cell(row=last_row, column=COL_D).value

        lane_info.append((lane_name, first_g, first_h, summary_d))

        # 数据行范围 (不含汇总行)
        data_rows = rows[:-1] if len(rows) > 1 else rows
        data_first = data_rows[0]
        data_last = data_rows[-1]

        # 写入 Lane编号到数据行 A列
        for row in data_rows:
            cell = ws.cell(row=row, column=COL_A)
            cell.value = lane_name
            cell.alignment = CENTER

        # 合并数据行 A列
        if len(data_rows) > 1:
            ws.merge_cells(
                start_row=data_first, start_column=COL_A,
                end_row=data_last, end_column=COL_A
            )

        # 仅数据行 A列~K列 标框线（汇总行不加）
        for row in data_rows:
            for col in range(COL_A, COL_K + 1):
                cell = ws.cell(row=row, column=col)
                cell.border = ALL_BORDER

        # 汇总行不加框线、不合并、不写Lane

    print(f'  Sheet {name}: {len(groups)} 组, Lane={name}1~{name}{len(groups)}')
    return lane_info


def fill_huanhua_sheet(ws_hh, all_lane_info):
    """填充文库环化工作表"""
    # 清空数据行（表头行1不动, clear_pooling已保留原始表头）
    if ws_hh.max_row > 1:
        ws_hh.delete_rows(2, ws_hh.max_row - 1)

    for i, (lane, g_val, h_val, d_val) in enumerate(all_lane_info):
        row = i + 2

        # A列: Lane编号
        c = ws_hh.cell(row=row, column=1)
        c.value = lane
        c.alignment = CENTER

        # B列, C列, L列: 留空(用户手工填)

        # D列: =C*0.66*1
        c = ws_hh.cell(row=row, column=4)
        c.value = f'=C{row}*0.66*1'
        c.alignment = CENTER

        # E列: 数据量
        c = ws_hh.cell(row=row, column=5)
        c.value = d_val
        c.alignment = CENTER

        # F列: =D
        c = ws_hh.cell(row=row, column=6)
        c.value = f'=D{row}'
        c.alignment = CENTER

        # G列: =F/B
        c = ws_hh.cell(row=row, column=7)
        c.value = f'=F{row}/B{row}'
        c.alignment = CENTER

        # H列: 文库类型
        c = ws_hh.cell(row=row, column=8)
        c.value = g_val
        c.alignment = CENTER

        # I列: 客户单位
        c = ws_hh.cell(row=row, column=9)
        c.value = h_val
        c.alignment = CENTER

        # M列(13): =22*L/F
        c = ws_hh.cell(row=row, column=13)
        c.value = f'=22*L{row}/F{row}'
        c.alignment = CENTER

        # 框线 A~K
        for col in range(1, 12):
            ws_hh.cell(row=row, column=col).border = ALL_BORDER

    print(f'  文库环化: 写入 {len(all_lane_info)} 行')


def main():
    wb = openpyxl.load_workbook(DST)

    all_lane_info = []

    for name in SHEETS:
        print(f'\n{"─"*50}')
        print(f'处理工作表 [{name}]')
        print(f'{"─"*50}')
        info = process_abc_sheet(wb[name], name)
        all_lane_info.extend(info)

    print(f'\n{"─"*50}')
    print(f'填充 [文库环化]')
    print(f'{"─"*50}')
    fill_huanhua_sheet(wb['文库环化'], all_lane_info)

    wb.save(DST)
    print(f'\n{"="*50}')
    print(f'✅ 步骤五完成 → {DST}')


if __name__ == '__main__':
    main()
