# -*- coding: utf-8 -*-
"""指导书 docx 格式规范化（只改格式，不动文字）。

规则：
- 正文 Normal(14pt)   : 宋体 + Times New Roman，14pt，两端对齐，首行缩进 2 字符
- 空段落              : 统一字体，不加缩进
- 封面大标题(42pt)    : 黑体 42pt 居中，补中文字体
- 副标题(15pt)        : 宋体 15pt，补中文字体
- Heading 2(章)       : 黑体 16pt 加粗 居中
- Heading 3/4/5       : 黑体 14pt 加粗 左对齐
- HTML Preformatted   : Consolas + 宋体，左对齐，无缩进，保持字号
输出到新文件，不覆盖原文件。
"""
import sys
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH

SRC = '配置核查作业指导书_v2.1草稿.docx'
DST = '配置核查作业指导书_v2.1草稿_格式规范化.docx'

CN_BODY = '宋体'
CN_HEAD = '黑体'
EN_BODY = 'Times New Roman'
EN_MONO = 'Consolas'


def set_run_font(run, east, ascii_font=EN_BODY, size=None, bold=None):
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:eastAsia'), east)
    rFonts.set(qn('w:ascii'), ascii_font)
    rFonts.set(qn('w:hAnsi'), ascii_font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold


def clear_indent(p):
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    # 清除左缩进、悬挂缩进
    for attr in ('w:left', 'w:hanging', 'w:leftChars', 'w:hangingChars'):
        if ind.get(qn(attr)) is not None:
            del ind.attrib[qn(attr)]


def set_first_line_chars(p, chars=200):
    """首行缩进 chars 个字符（firstLineChars 自适应字号），并兜底设 firstLine。"""
    pPr = p._p.get_or_add_pPr()
    ind = pPr.find(qn('w:ind'))
    if ind is None:
        ind = OxmlElement('w:ind')
        pPr.append(ind)
    ind.set(qn('w:firstLineChars'), str(chars))
    ind.set(qn('w:firstLine'), str(int(chars * 240)))  # 兜底：每字符约 240 twips


def para_lead_size(p):
    """段落首个非空 run 的显式字号(pt)，无则 None。"""
    for r in p.runs:
        if r.text.strip():
            if r.font.size is not None:
                return r.font.size.pt
            return None
    return None


doc = Document(SRC)

for p in doc.paragraphs:
    style = p.style.name
    text = p.text

    if style == 'HTML Preformatted':
        # 代码块：等宽西文 + 中文宋体，左对齐，无缩进，保持字号
        clear_indent(p)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = None
        for r in p.runs:
            set_run_font(r, CN_BODY, EN_MONO)
        continue

    if style.startswith('Heading'):
        # 标题统一黑体加粗
        lvl = int(style.split()[-1])
        if lvl == 2:
            size, align = 16, WD_ALIGN_PARAGRAPH.CENTER
        else:
            size, align = 14, WD_ALIGN_PARAGRAPH.LEFT
        clear_indent(p)
        p.paragraph_format.alignment = align
        p.paragraph_format.first_line_indent = None
        for r in p.runs:
            set_run_font(r, CN_HEAD, EN_BODY, size=size, bold=True)
        continue

    # Normal / Normal (Web) 等正文
    lead = para_lead_size(p)
    if lead is None or lead == 14.0:
        # 正文：宋体+TNR 14pt，两端对齐，首行缩进 2 字符（空段不加缩进）
        clear_indent(p)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        if text.strip():
            set_first_line_chars(p, 200)
        else:
            p.paragraph_format.first_line_indent = None
        for r in p.runs:
            set_run_font(r, CN_BODY, EN_BODY, size=14.0)
    elif lead == 42.0:
        # 封面大标题：黑体 42pt 居中
        clear_indent(p)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = None
        for r in p.runs:
            set_run_font(r, CN_HEAD, EN_BODY, size=42.0, bold=True)
    elif lead == 15.0:
        # 副标题：宋体 15pt
        clear_indent(p)
        p.paragraph_format.first_line_indent = None
        for r in p.runs:
            set_run_font(r, CN_BODY, EN_BODY, size=15.0)
    else:
        # 其他特殊字号：只补中文字体，不改字号/对齐/缩进
        for r in p.runs:
            set_run_font(r, CN_BODY, EN_BODY)

doc.save(DST)
print('已生成：' + DST)
