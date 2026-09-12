# ⚠️ 已停用（勿再运行）：本脚本会改动目录/分节结构，用户明确要求「不要动目录」，改动已还原。
#    保留仅作审计留痕；如确需重做页码三段式，须先取得用户同意，并在执行前备份。
# -*- coding: utf-8 -*-
"""把封面标题与“目录”标题从目录条目里摘掉。

原因：
  1) 封面「配置核查 / 作业指导书」两个段落本身是正文样式，但文字套了字符样式
     Heading1Char（Heading 1 的字符样式），目录域带 \\u（按大纲级别收集）时被当成一级条目收进去；
  2) 「目录」这一行用的是 Heading 2 段落样式，\\o "1-3" 自然也会收它。

做法：
  - 去掉封面标题 run 上的 Heading*Char 字符样式，同时补 <w:b/> 保留加粗外观（原加粗来自该字符样式）
  - 「目录」段落由 Heading2 改为正文样式，用直接格式复刻原外观（居中、黑体、三号、加粗）

用法：python _fix_guide_toc_entries.py [--apply]
"""
import re
import sys

import docx
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

PATH = "配置核查作业指导书_v2.2.docx"
APPLY = "--apply" in sys.argv
CHAR_RE = re.compile(r"^Heading\d+Char$", re.I)


def strip_char_style(p_el, ensure_bold=True):
    """去掉段落内所有 Heading*Char 字符样式引用；顺带补加粗，避免外观变化。

    w:rPr 子元素有固定顺序（rStyle, rFonts, b, i, ... sz ...），w:b 必须插在 rFonts 之后。
    """
    n = 0
    for rs in list(p_el.iter(qn("w:rStyle"))):
        val = rs.get(qn("w:val")) or ""
        if CHAR_RE.match(val):
            rPr = rs.getparent()
            rPr.remove(rs)
            n += 1
            if ensure_bold and rPr.find(qn("w:b")) is None:
                b = rPr.makeelement(qn("w:b"), {})
                rf = rPr.find(qn("w:rFonts"))
                if rf is not None:
                    rf.addnext(b)
                else:
                    rPr.insert(0, b)
    return n


def main():
    d = docx.Document(PATH)
    body = d.element.body

    # ---- 1) 封面标题：找“没有任何 pStyle（或非 Heading）”但含 Heading*Char 的段落 ----
    fixed_cover = []
    for p_el in body.iter(qn("w:p")):
        # 跳过真正的标题段落
        pPr = p_el.find(qn("w:pPr"))
        st = None
        if pPr is not None:
            st_el = pPr.find(qn("w:pStyle"))
            st = st_el.get(qn("w:val")) if st_el is not None else None
        if st and st.lower().startswith("heading"):
            continue
        txt = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()
        styles = [rs.get(qn("w:val")) for rs in p_el.iter(qn("w:rStyle"))]
        if any(CHAR_RE.match(s or "") for s in styles):
            fixed_cover.append((txt[:24], styles))
            if APPLY:
                strip_char_style(p_el)
    print("=== 1. 封面标题去掉 Heading*Char 字符样式 ===")
    for t, s in fixed_cover:
        print("   %-16r 原字符样式=%s" % (t, s))
    if not fixed_cover:
        print("   （无）")

    # ---- 2) “目录”标题：Heading2 → 正文样式 + 直接格式复刻 ----
    fixed_toc = None
    for p_el in body.iter(qn("w:p")):
        txt = "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()
        if txt != "目录":
            continue
        pPr = p_el.find(qn("w:pPr"))
        st_el = pPr.find(qn("w:pStyle")) if pPr is not None else None
        st = st_el.get(qn("w:val")) if st_el is not None else None
        if st and st.lower().startswith("heading"):
            fixed_toc = st
            if APPLY:
                pPr.remove(st_el)
    print("\n=== 2. “目录”标题由标题样式改为正文 ===")
    print("   原样式:", fixed_toc if fixed_toc else "（未找到 Heading 样式的“目录”段）")

    if not APPLY:
        print("\n（干跑，未写文件；加 --apply 执行）")
        return 0

    d.save(PATH)
    print("\n已写入", PATH)

    # 用 python-docx 复核：目录段落是否还在标题集合里
    d2 = docx.Document(PATH)
    hs = [p.text.strip() for p in d2.paragraphs if p.style.name.startswith("Heading")]
    print("标题段落数:", len(hs), "| 其中含“目录”条目:", "目录" in hs)
    return 0


if __name__ == "__main__":
    sys.exit(main())
