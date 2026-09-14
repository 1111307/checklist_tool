# -*- coding: utf-8 -*-
r"""指导书三段式分节（封面无页码 → 目录罗马 → 正文阿拉伯从 1 起）+ 页眉。

用法：
  python _guide_sections.py --apply   # 需先备份
"""
import copy
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PATH = "配置核查作业指导书_v2.2.docx"
TITLE = "配置核查作业指导书"

sys.path.insert(0, ".")
from md2docx import _set_pgnum, _set_doc_header, _set_section_footer_page  # noqa: E402


def insert_section_break(body, sentinel, before_el):
    """在 before_el 之前插入一个携带 sectPr 副本的空段（节分隔），副本清掉页眉页脚引用与页码设置。"""
    sect = copy.deepcopy(sentinel)
    for tag in ('w:headerReference', 'w:footerReference', 'w:pgNumType', 'w:titlePg'):
        for e in sect.findall(qn(tag)):
            sect.remove(e)
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    pPr.append(sect)
    p.append(pPr)
    before_el.addprevious(p)


def main():
    d = Document(PATH)
    body = d.element.body
    sdt = body.find(qn('w:sdt'))
    sentinel = body.find(qn('w:sectPr'))
    assert sdt is not None and sentinel is not None, "未找到目录 sdt 或 body sectPr"
    if len(d.sections) != 1:
        print("!! 已分节（%d 节），跳过插入" % len(d.sections))
    else:
        insert_section_break(body, sentinel, sdt)      # 封面节结束
        insert_section_break(body, sentinel, sdt.getnext())  # 目录节结束
        d.save(PATH)
        print("已插入 2 个分节符")
        d = Document(PATH)
    print("节数:", len(d.sections))
    if len(d.sections) != 3:
        print("!! 节数不为 3，中止")
        return 1
    sec_cover, sec_toc, sec_body = d.sections
    # 目录节：罗马页码从 I 起 + 页眉 + 页码
    _set_pgnum(sec_toc, 'upperRoman', 1)
    _set_section_footer_page(sec_toc)
    _set_doc_header(sec_toc, TITLE)
    # 正文节：阿拉伯页码从 1 起 + 页眉 + 页码
    _set_pgnum(sec_body, 'decimal', 1)
    _set_section_footer_page(sec_body, arabic_switch=True)
    _set_doc_header(sec_body, TITLE)
    d.save(PATH)
    print("已配置三节页眉页码，保存")
    return 0


if __name__ == "__main__":
    sys.exit(main())
