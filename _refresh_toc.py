# ⚠️ 已停用（勿再运行）：本脚本会改动目录/分节结构，用户明确要求「不要动目录」，改动已还原。
#    保留仅作审计留痕；如确需重做页码三段式，须先取得用户同意，并在执行前备份。
# -*- coding: utf-8 -*-
"""用 Word 打开指导书：更新目录域与页码域 → 保存 → 导出前若干页 PDF 供核对。

用法：python _refresh_toc.py [导出页数，默认 4]
"""
import os
import sys
import time

import win32com.client as win32

PATH = os.path.abspath("配置核查作业指导书_v2.2.docx")
PDF = os.path.abspath("_check_前几页.pdf")
N = int(sys.argv[1]) if len(sys.argv) > 1 else 4

word = win32.DispatchEx("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
doc = None
try:
    print("打开文档 ...")
    doc = word.Documents.Open(PATH)
    print("段落数:", doc.Paragraphs.Count, "节数:", doc.Sections.Count)

    # 更新目录
    n_toc = doc.TablesOfContents.Count
    print("目录数量:", n_toc)
    for i in range(1, n_toc + 1):
        doc.TablesOfContents(i).Update()
        print("  已更新目录", i)

    # 更新全部域（页码等）
    doc.Fields.Update()
    print("已更新全部域")

    # 打印前 3 节页码设置
    for i in range(1, doc.Sections.Count + 1):
        s = doc.Sections(i)
        pn = s.Footers(1).PageNumbers
        print("  节%d: 起始页码=%s 页码类型=%s 页脚文本=%r" % (
            i, pn.StartingNumber, pn.NumberStyle,
            s.Footers(1).Range.Text.strip()[:20]))

    doc.Save()
    print("已保存")

    # 导出前 N 页
    if os.path.exists(PDF):
        os.remove(PDF)
    doc.ExportAsFixedFormat(PDF, 17, False, 0, 3, 1, N, 0, True, True, 0, True, True, False)
    print("已导出前 %d 页 ->" % N, PDF)
finally:
    if doc is not None:
        doc.Close(False)
    word.Quit()
    time.sleep(1)
print("完成")
