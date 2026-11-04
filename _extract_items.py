# -*- coding: utf-8 -*-
"""从指导书 v2.0.0 提取 136 个核查项 → checklist_items.json（人工核查台数据源）。"""
import json
import re
from docx import Document
from docx.oxml.ns import qn

doc = Document(r'配置核查作业指导书_v2.0.0.docx')

# 样式识别：旧版 docx 的 styleId 是 '2'~'6'，现版已规范化为 Heading1~5，
# 两套都要认，否则在新文档上提取结果为 0 项（历史上踩过）。
STYLE_LVL = {'2': 1, '3': 2, '4': 3, '5': 4, '6': 5,
             'Heading1': 1, 'Heading2': 2, 'Heading3': 3,
             'Heading4': 4, 'Heading5': 5}


def get_lvl(p):
    pPr = p._p.pPr
    if pPr is None: return None
    ps = pPr.find(qn('w:pStyle'))
    if ps is None: return None
    val = ps.get(qn('w:val'))
    if val in STYLE_LVL:
        return STYLE_LVL[val]
    m = re.match(r'Heading\s*(\d)', p.style.name or '')
    return int(m.group(1)) if m else None

items = []
cur_ch = None
for p in doc.paragraphs:
    lv = get_lvl(p)
    t = p.text.strip()
    if lv == 1:
        cur_ch = t
    elif lv == 2:
        m = re.match(r'(\d+\.\d+)\s+(.*)', t)
        if m:
            items.append({'ch': cur_ch, 'id': m.group(1), 'title': m.group(2)})

# 手动核查的"参考指导书"列 = 章名 + 编号 + 标题
for it in items:
    it['guide'] = f"《配置核查作业指导书v2.0.0》{it['ch']} {it['id']}：{it['title']}"
    it['cat'] = re.sub(r'^\d+\s*', '', it['ch'])  # 章名去掉编号作为类别

with open('checklist_items.json', 'w', encoding='utf-8') as f:
    json.dump(items, f, ensure_ascii=False, indent=1)

from collections import Counter
print(f'提取 {len(items)} 项')
print(dict(Counter(i['ch'] for i in items)))
