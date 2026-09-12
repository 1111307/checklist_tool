# ⚠️ 已停用（勿再运行）：本脚本会改动目录/分节结构，用户明确要求「不要动目录」，改动已还原。
#    保留仅作审计留痕；如确需重做页码三段式，须先取得用户同意，并在执行前备份。
# -*- coding: utf-8 -*-
"""修正分节结构：复用目录内容控件里原有的分节符。

背景：目录（TOC）被包在内容控件 sdt 里，而该 sdt 的最后一段本来就带 sectPr，
      是 Word 里真实存在的分节符；python-docx 的 sections 不进入 sdt，所以此前误判为「单节」。
      上一版脚本在 sdt 之后又插了一个带 sectPr 的空段落，造成 4 个 sectPr、页码落在错误的节上。

本脚本：
  1. 删除上一版插入的空段落（多余的第三节）
  2. 目录内原有的 sectPr 设为 lowerRoman 从 1 起，并挂上与正文相同的页脚（- PAGE -）
  3. 保留封面节（无页脚）与正文节（decimal 从 1 起）

用法：python _fix_guide_sections.py [--apply]
"""
import copy
import re
import sys

import docx
from docx.oxml.ns import qn

PATH = "配置核查作业指导书_v2.2.docx"
APPLY = "--apply" in sys.argv


def sect_xml(el, keep_ns=True):
    x = re.sub(r"\s+", " ", el.xml)
    if not keep_ns:
        x = re.sub(r'xmlns:\w+="[^"]*"\s*', "", x)
    return x


def main():
    d = docx.Document(PATH)
    body = d.element.body

    sects = list(body.iter(qn("w:sectPr")))
    print("文档中共有 sectPr: %d" % len(sects))
    for i, s in enumerate(sects):
        print("  [%d] %s" % (i, re.sub(r'xmlns:\w+="[^"]*"\s*', "", re.sub(r"\s+", " ", s.xml))[:190]))

    if len(sects) != 4:
        print("\n预期 4 个（封面/目录/多余/正文），实际 %d 个，请人工确认后再运行。" % len(sects))
        return 1

    cover, toc, extra, bodysect = sects

    # 1) 删掉多余的那一节（sectPr 的父级是 pPr，再往上是段落 p，整段移除）
    _pPr = extra.getparent()
    _para = _pPr.getparent() if _pPr is not None else None
    if _para is not None and _para.tag == qn("w:p"):
        print("\n将删除多余段落（含第 3 个 sectPr），该段 w:t=%r"
              % [t.text for t in _para.iter(qn("w:t"))])
        if APPLY:
            _para.getparent().remove(_para)
    else:
        print("\n!! 未定位到多余段落，跳过删除")

    # 2) 目录节：页码 lowerRoman 从 1 起 + 挂正文同款页脚引用
    for el in toc.findall(qn("w:pgNumType")):
        toc.remove(el)
    pn = toc.makeelement(qn("w:pgNumType"), {})
    pn.set(qn("w:fmt"), "lowerRoman")
    pn.set(qn("w:start"), "1")
    toc.append(pn)
    if toc.get(qn("w:type")) is None:
        toc.set(qn("w:type"), "nextPage")

    body_frefs = bodysect.findall(qn("w:footerReference"))
    for el in toc.findall(qn("w:footerReference")):
        toc.remove(el)
    for ref in body_frefs:
        if ref.get(qn("w:type")) == "default":
            toc.append(copy.deepcopy(ref))
            break
    print("目录节 -> lowerRoman 从 1 起，页脚引用 %d 个" % len(toc.findall(qn("w:footerReference"))))

    # 3) 封面节：确保无页脚引用、无 pgNumType、下一页分节
    for el in list(cover.findall(qn("w:footerReference"))) + list(cover.findall(qn("w:pgNumType"))):
        cover.remove(el)
    if cover.get(qn("w:type")) is None:
        cover.set(qn("w:type"), "nextPage")
    print("封面节 -> 无页脚引用（不显示页码）")

    # 4) 正文节：decimal 从 1 起
    for el in bodysect.findall(qn("w:pgNumType")):
        bodysect.remove(el)
    pn3 = bodysect.makeelement(qn("w:pgNumType"), {})
    pn3.set(qn("w:fmt"), "decimal")
    pn3.set(qn("w:start"), "1")
    bodysect.append(pn3)
    print("正文节 -> decimal 从 1 起")

    if not APPLY:
        print("\n（干跑，未写文件；加 --apply 执行）")
        return 0

    d.save(PATH)
    print("\n已写入", PATH)

    d2 = docx.Document(PATH)
    sects2 = list(d2.element.body.iter(qn("w:sectPr")))
    print("=== 校验：现有 sectPr %d 个 ===" % len(sects2))
    for i, s in enumerate(sects2):
        sx = re.sub(r"\s+", " ", s.xml)
        print("  [%d] pgNum=%s 页脚引用=%d" % (
            i, re.findall(r'<w:pgNumType[^>]*/>', sx),
            len(re.findall("footerReference", sx))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
