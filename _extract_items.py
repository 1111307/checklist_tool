# -*- coding: utf-8 -*-
"""从指导书 v2.2 提取 113 个核查项 → checklist_items.json（手动核查界面数据源）。"""
import json
import re
from docx import Document
from docx.oxml.ns import qn

doc = Document(r'配置核查作业指导书_v2.2.docx')

def get_lvl(p):
    pPr = p._p.pPr
    if pPr is None: return None
    ps = pPr.find(qn('w:pStyle'))
    if ps is None: return None
    return {'2':1,'3':2,'4':3,'5':4,'6':5}.get(ps.get(qn('w:val')))

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
    it['guide'] = f"《配置核查作业指导书v2.2》{it['ch']} {it['id']}：{it['title']}"
    it['cat'] = re.sub(r'^\d+\s*', '', it['ch'])  # 章名去掉编号作为类别

with open('checklist_items.json', 'w', encoding='utf-8') as f:
    json.dump(items, f, ensure_ascii=False, indent=1)

from collections import Counter
print(f'提取 {len(items)} 项')
print(dict(Counter(i['ch'] for i in items)))
