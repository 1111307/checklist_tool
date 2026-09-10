# -*- coding: utf-8 -*-
"""把 _matrix_from_guide.json 的 136 项对象覆盖集写入核查表 C~M 列（√ / —）。

行布局（结构改动后的固定映射）：
  第1章  R4-R30   ← 1.1-1.27
  第3章  R31-R44  ← 3.1-3.14
  第4章  R45-R77  ← 4.1-4.33
  第2章  R78-R93  ← 2.1-2.16
  第5章  R94-R116 ← 5.1-5.23
  第6章  R117-R125 / 第7章 R126-R130 / 第8章 R131-R134 / 第9章 R135-R138 / 第10章 R139
列：C~M 共 11 列 = COLS 顺序（华为交换机插在 Tomcat 之后）。
"""
import json
import openpyxl
from openpyxl.styles import Alignment, Font

XLSX = '配置核查表_v2.0.0_标注自动验证.xlsx'
COLS = ['Win7', 'WinXP', '中标麒麟', '银河麒麟', 'Nginx', 'Tomcat', '华为交换机', 'Mysql', 'SQLSever', '达梦', 'redis']
FIRST_COL = 3                      # C
GROUP_START = {'1': 4, '3': 31, '4': 45, '2': 78, '5': 94, '6': 117, '7': 126, '8': 131, '9': 135, '10': 139}
XLSX_NAME = {'1': '系统安全', '2': '用户安全', '3': '数据安全', '4': '应用安全',
             '5': '网络安全', '6': '物理安全', '7': '组织机构', '8': '规章制度',
             '9': '管理实施', '10': '协议安全审计'}

matrix = json.load(open('_matrix_from_guide.json', encoding='utf-8'))

# 逐章内按编号排序，与行号一一对应
row_map = {}
for ch, start in GROUP_START.items():
    keys = sorted([k for k in matrix if k.split('.')[0] == ch], key=lambda x: int(x.split('.')[1]))
    for i, k in enumerate(keys):
        row_map[start + i] = k

wb = openpyxl.load_workbook(XLSX)
ws = wb['Sheet1']

# 写前断言：行数与章名对应
bad = [r for r, k in row_map.items() if XLSX_NAME[k.split('.')[0]] not in matrix[k]['ch']]
assert not bad, f'章节归属异常: {bad}'
# 行号必须落在数据区且连续
rows = sorted(row_map)
assert rows == list(range(4, 140)), f'行覆盖不连续: {len(rows)} 行'
assert len(rows) == 136, len(rows)

font = Font(name='宋体', size=10)
center = Alignment(horizontal='center', vertical='center')

# 表头对象名统一（修正原始表 "Tomocat" 拼写笔误，与指导书/脚本口径一致）
HEADER = ['Windows7', 'WindowsXP', '中标麒麟', '银河麒麟', 'Nginx', 'Tomcat',
          '华为交换机', 'Mysql', 'SQLSever', '达梦', 'redis']
for j, name in enumerate(HEADER):
    c = ws.cell(3, FIRST_COL + j)
    c.value = name
    c.font = font
    c.alignment = center

written = 0
for r, k in row_map.items():
    objs = set(matrix[k]['objs'])
    for j, col in enumerate(COLS):
        c = ws.cell(r, FIRST_COL + j)
        c.value = '√' if col in objs else '—'
        c.font = font
        c.alignment = center
        written += 1

wb.save(XLSX)

# ---- 写后回读校验 ----
wb2 = openpyxl.load_workbook(XLSX)
ws2 = wb2['Sheet1']
errs = []
for r, k in row_map.items():
    got = []
    for j, col in enumerate(COLS):
        v = ws2.cell(r, FIRST_COL + j).value
        if v == '√':
            got.append(col)
        elif v != '—':
            errs.append(f'R{r} {col} 值异常: {v!r}')
    if got != [c for c in COLS if c in set(matrix[k]['objs'])]:
        errs.append(f'R{r} {k} 读回不一致: {got} != {matrix[k]["objs"]}')
assert not errs, '\n'.join(errs[:20])
# 表头对象名必须与 COLS 一致
hdr = [ws2.cell(3, FIRST_COL + j).value for j in range(11)]
assert hdr == HEADER, hdr
assert ws2.cell(2, 14).value == '工具自动验证', ws2.cell(2, 14).value

print(f'已写入 {written} 个单元格（{len(row_map)} 行 × 11 列），回读校验通过')
print(f'表头: {hdr}')
py = sum(1 for r, k in row_map.items() if '华为交换机' in matrix[k]['objs'])
print(f'勾选“华为交换机”的行数: {py}（第5章 23 项 + 3.6/4.4/4.5/4.18 等）')
print('抽查：')
for r in (4, 30, 31, 78, 94, 116, 139):
    k = row_map[r]
    print(f'  R{r} {k:5s} {matrix[k]["title"][:22]:24s} ' +
          ''.join('√' if c in set(matrix[k]['objs']) else '—' for c in COLS))
