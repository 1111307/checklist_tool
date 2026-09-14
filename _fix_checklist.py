# -*- coding: utf-8 -*-
r"""核查表 v2.0.0：填写说明补「（4）网络设备：华为交换机」，末列「工具自动验证」按三态填色。

以 Desktop 实物包 140×14 版为源，统一写回四处副本：
  Desktop\checklist_tool\配置核查表_v2.0.0.xlsx      （源）
  仓库根 配置核查表_v2.0.0.xlsx                       （原为过期 139×13，一并统一）
  仓库根 配置核查表_v2.0.0_标注自动验证.xlsx
  仓库 checklist_tool\配置核查表_v2.0.0_标注自动验证.xlsx
"""
import os
import shutil

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

SRC = r"C:/Users/ryan.xiong/Desktop/checklist_tool/配置核查表_v2.0.0.xlsx"
TARGETS = [
    SRC,
    "配置核查表_v2.0.0.xlsx",
    "配置核查表_v2.0.0_标注自动验证.xlsx",
    "checklist_tool/配置核查表_v2.0.0_标注自动验证.xlsx",
]

FILLS = {
    "可自动化":   ("C6EFCE", "006100"),
    "部分可自动化": ("FFEB9C", "9C6500"),
    "需人工":     ("FFC7CE", "9C0006"),
}
NEW_LINE = "(4)网络设备：华为交换机"

wb = openpyxl.load_workbook(SRC)
ws = wb.worksheets[0]
LAST = ws.max_column

# ---- 1. A1 说明补第（4）项 ----
a1 = ws.cell(1, 1).value or ""
if NEW_LINE not in a1:
    lines = a1.split("\n")
    out, inserted = [], False
    for ln in lines:
        if (not inserted) and ln.strip().startswith("(3)"):
            out.append(ln)
            out.append(NEW_LINE)
            inserted = True
        else:
            out.append(ln)
    if not inserted:          # 兜底：追加到末尾
        out.append(NEW_LINE)
    ws.cell(1, 1).value = "\n".join(out)
    print("A1 已补第（4）项")
ws.cell(1, 1).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
h = ws.row_dimensions[1].height
if h and h < 200:
    ws.row_dimensions[1].height = h + 18      # 多一行，适当加高

# ---- 2. 末列三态填色 ----
stat = {}
for r in range(4, ws.max_row + 1):
    c = ws.cell(r, LAST)
    v = str(c.value or "").strip()
    if not v:
        continue
    for key, (bg, fg) in FILLS.items():
        if v.startswith(key):
            c.fill = PatternFill("solid", start_color=bg, end_color=bg)
            c.font = Font(name=c.font.name, size=c.font.size, color=fg, bold=False)
            stat[key] = stat.get(key, 0) + 1
            break
print("末列填色统计:", stat, "（列 %d，数据行 4~%d）" % (LAST, ws.max_row))

# ---- 3. 写回各副本 ----
for t in TARGETS:
    d = os.path.dirname(t)
    if d and not os.path.isdir(d):
        print("!! 目录不存在，跳过:", t); continue
    wb.save(t)
    print("已保存:", t)

# ---- 4. 复核 ----
chk = openpyxl.load_workbook(SRC).worksheets[0]
print("\n复核 A1 末两行:", chk.cell(1, 1).value.split("\n")[-3:])
for r in (4, 5, 60):
    c = chk.cell(r, chk.max_column)
    print("  r%d 末列: %-6s 填充=%s 字色=%s" % (
        r, str(c.value)[:6], c.fill.start_color.rgb, c.font.color.rgb if c.font.color else None))
