# -*- coding: utf-8 -*-
"""Markdown -> docx 转换器（纯 python-docx，无需 pandoc）。

用法：python md2docx.py <输入.md> <输出.docx>
支持：标题(#/##/###)、表格、无序/有序列表、段落、**加粗**、图片 ![图注](路径)。
字体按论文规格：中文正文宋体（小四 12pt）、标题黑体（三号/四号/小四）、
西文与数字 Times New Roman；表格与图注五号 10.5pt；正文行距 1.5、首行缩进 2 字符。
"""
import re, sys, os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

IMG_WIDTH = Cm(15.5)
IMG_MAX_HEIGHT = 12.5

try:
    from PIL import Image as _PILImage
except ImportError:
    _PILImage = None

EAST_BODY = '宋体'
EAST_HEAD = '黑体'
LATIN = 'Times New Roman'

BODY_SIZE = 12.0     # 正文小四
TABLE_SIZE = 10.5    # 表格五号
CAPTION_SIZE = 10.5  # 图注五号
NOTE_SIZE = 9.5      # 图下说明（小五）
HEAD_SIZES = {1: 16.0, 2: 14.0, 3: 12.0}   # 三号 / 四号 / 小四

def _set_font(run, east, latin=LATIN, size=BODY_SIZE, bold=False, color=None):
    run.font.name = latin
    run._element.rPr.rFonts.set(qn('w:eastAsia'), east)
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color

MONO_LATIN = 'Consolas'
CODE_SIZE = 10.5    # 代码块/行内代码字号（五号）


def _add_runs(p, text, east=EAST_BODY, size=BODY_SIZE, bold=False):
    """处理 **加粗** 与 `行内代码`：代码用等宽字体，去掉反引号。"""
    for part in re.split(r'(\*\*.*?\*\*)', text):
        if not part:
            continue
        b = part.startswith('**') and part.endswith('**')
        body = part[2:-2] if b else part
        for seg in re.split(r'(`[^`]*`)', body):
            if not seg:
                continue
            if seg.startswith('`') and seg.endswith('`') and len(seg) > 2:
                _set_font(p.add_run(seg[1:-1]), east, latin=MONO_LATIN,
                          size=min(size, CODE_SIZE + 0.5), bold=bold or b)
            else:
                _set_font(p.add_run(seg), east, size=size, bold=bold or b)

def _first_line_indent_2chars(p):
    """首行缩进 2 字符：用 w:firstLineChars（随字号自适应），并给 firstLine 兜底值。"""
    ind = p._p.get_or_add_pPr().get_or_add_ind()
    ind.set(qn('w:firstLineChars'), '200')
    ind.set(qn('w:firstLine'), str(int(BODY_SIZE * 2 * 20)))


def _heading(doc, text, level):
    p = doc.add_heading(level=level)
    _add_runs(p, text, east=EAST_HEAD, size=HEAD_SIZES.get(level, 12.0), bold=True)
    for r in p.runs:
        r.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_before = Pt(12 if level <= 2 else 8)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    return p

def _set_row_flags(row, header=False):
    """行内禁止跨页断裂；首行作为表头逐页重复。"""
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement('w:cantSplit'))
    if header:
        trPr.append(OxmlElement('w:tblHeader'))


def _add_footer_page_number(doc, size=CAPTION_SIZE):
    for section in doc.sections:
        p = section.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
        instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
        end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
        run._r.append(begin); run._r.append(instr); run._r.append(end)
        _set_font(run, EAST_BODY, size=size)


def _table(doc, rows):
    rows = [[c.strip() for c in r.strip().strip('|').split('|')] for r in rows]
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = 'Table Grid'
    t.autofit = True
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            c = t.cell(i, j)
            p = c.paragraphs[0]
            _add_runs(p, cell, size=TABLE_SIZE, bold=(i == 0))
        _set_row_flags(t.rows[i], header=(i == 0))
        if i == 0:
            for c in t.rows[0].cells:
                for p in c.paragraphs:
                    p.paragraph_format.keep_with_next = True
    return t

def convert(md_path, out_path):
    doc = Document()
    # 默认样式
    normal = doc.styles['Normal']
    normal.font.name = LATIN
    normal._element.rPr.rFonts.set(qn('w:eastAsia'), EAST_BODY)
    normal.font.size = Pt(BODY_SIZE)

    lines = open(md_path, encoding='utf-8').read().splitlines()
    i = 0
    n = len(lines)
    while i < n:
        ln = lines[i].rstrip()
        if ln.strip() == '':
            i += 1
            continue
        # 标题
        m = re.match(r'^(#{1,6})\s+(.*)', ln)
        if m:
            _heading(doc, m.group(2).strip(), len(m.group(1)))
            i += 1
            continue
        # 围栏代码块 ``` ... ```（逐行输出等宽字体，无首行缩进）
        if ln.strip().startswith('```'):
            i += 1
            code_lines = []
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i].rstrip())
                i += 1
            i += 1                      # 跳过闭合围栏
            while code_lines and not code_lines[0].strip():
                code_lines.pop(0)
            while code_lines and not code_lines[-1].strip():
                code_lines.pop()
            for k, cl in enumerate(code_lines):
                cp = doc.add_paragraph()
                cp.paragraph_format.first_line_indent = Pt(0)
                cp.paragraph_format.left_indent = Cm(0.6)
                cp.paragraph_format.line_spacing = 1.15
                cp.paragraph_format.space_before = Pt(4 if k == 0 else 0)
                cp.paragraph_format.space_after = Pt(4 if k == len(code_lines) - 1 else 0)
                cp.paragraph_format.keep_with_next = (k < len(code_lines) - 1)
                if cl.strip():
                    _set_font(cp.add_run(cl), EAST_BODY, latin=MONO_LATIN, size=CODE_SIZE)
            continue
        # 图片：![图注](相对/绝对路径)
        m = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)\s*$', ln.strip())
        if m:
            alt, src = m.group(1).strip(), m.group(2).strip()
            path = src if os.path.isabs(src) else os.path.join(os.path.dirname(os.path.abspath(md_path)), src)
            if os.path.exists(path):
                pic = doc.add_paragraph()
                pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                pic.paragraph_format.first_line_indent = Pt(0)
                pic.paragraph_format.space_before = Pt(6)
                pic.paragraph_format.keep_with_next = True
                w_cm = 15.5
                if _PILImage is not None:
                    try:
                        with _PILImage.open(path) as _im:
                            w_px, h_px = _im.size
                        _ratio = h_px / max(w_px, 1)
                        if w_cm * _ratio > 12.5:
                            w_cm = 12.5 / _ratio
                    except Exception:
                        pass
                pic.add_run().add_picture(path, width=Cm(w_cm))
                if alt:
                    cap = doc.add_paragraph()
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    cap.paragraph_format.keep_with_next = True
                    _add_runs(cap, alt, size=CAPTION_SIZE)
                print('  插图：' + (alt or src))
            else:
                print('  !! 图片不存在：' + path)
            i += 1
            continue
        # 表格
        if ln.lstrip().startswith('|'):
            tbl = [ln]
            while i + 1 < n and lines[i + 1].lstrip().startswith('|'):
                i += 1
                tbl.append(lines[i])
            # 去掉分隔行 |---|---|
            tbl = [r for r in tbl if not re.match(r'^\s*\|?[\s:|-]+\|?\s*$', r)]
            _table(doc, tbl)
            i += 1
            continue
        # 无序列表
        if re.match(r'^\s*[-*+]\s+', ln):
            p = doc.add_paragraph(style='List Bullet')
            _add_runs(p, re.sub(r'^\s*[-*+]\s+', '', ln))
            i += 1
            continue
        # 有序列表
        if re.match(r'^\s*\d+\.\s+', ln):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.74)
            _add_runs(p, ln.strip())
            i += 1
            continue
        # 图注/说明行（> 开头）：小字、左缩进、无首行缩进
        if ln.startswith('>'):
            p = doc.add_paragraph()
            _add_runs(p, ln.lstrip('> ').strip(), size=NOTE_SIZE)
            p.paragraph_format.left_indent = Cm(0.6)
            p.paragraph_format.first_line_indent = Pt(0)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(4)
            i += 1
            continue
        # 普通段落
        p = doc.add_paragraph()
        _add_runs(p, ln)
        _first_line_indent_2chars(p)
        p.paragraph_format.line_spacing = 1.5
        # 以冒号结尾的引导句（「…如下：」）与后文（多为图）保持同页，避免孤句留在页尾
        if ln.rstrip().endswith(('：', ':')):
            p.paragraph_format.keep_with_next = True
        i += 1

    _add_footer_page_number(doc)
    doc.save(out_path)

if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('用法：python md2docx.py <输入.md> <输出.docx>')
        sys.exit(1)
    convert(sys.argv[1], sys.argv[2])
    print('已生成：' + sys.argv[2])
