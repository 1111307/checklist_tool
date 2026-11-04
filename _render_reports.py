# -*- coding: utf-8 -*-
"""渲染两份报告：md -> docx -> Word 更新域 -> 保存（**不碰指导书**）。

为什么必须有这一步：md2docx.py 产出的 docx 目录域是**空的**（无缓存条目、
无 _Toc 书签、无 PAGEREF 页码），读者打开只看到空目录。git 里那两份交付
docx 的 app.xml 写着 Application=Microsoft Office Word、Pages=42/23，
是「被 Word 打开更新域后重存」的痕迹。所以改完 md 不能只跑 md2docx.py。

用法
----
    python _render_reports.py                 # 重建两份报告 + 更新域（默认）
    python _render_reports.py --only-fields   # 不重建，只对现有 docx 更新域
    python _render_reports.py --check         # 只读：页数 / 目录条目 / 内联图
    python _render_reports.py --pdf           # 另存 PDF 供版式核对
    python _render_reports.py --only 报告      # 只处理一份（报告|手册）
    python _render_reports.py --out-dir .tmp  # 渲染到别的目录（验证用，不覆盖交付件）

注意
----
- 与 _refresh_toc.py 的区别：那个是**指导书**专用，改分节/页码结构，已停用
  （用户明确要求「不要动目录」）。本脚本只更新这两份报告的域，不碰结构。
- 跑完看「页数有没有变」：页数纹丝不动，多半是生成步骤被跳过了（历史上真
  发生过——docstring 写了 md→docx，代码里只有 Word 那一步）。
- 所有 PDF/页数类判断以渲染结果为准，不要从 XML 属性推算。
"""
import argparse
import hashlib
import os
import re
import subprocess
import sys
import zipfile

try:
    sys.stdout.reconfigure(errors="replace")  # GBK 控制台遇到 ✓/↔ 不炸
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(ROOT, "测评报告")

# (key, 显示名, md 相对路径, docx 相对路径)
REPORTS = [
    ("报告", "交叉验证报告", "测评报告/指导书与核查工具交叉验证报告.md",
     "测评报告/指导书与核查工具交叉验证报告.docx"),
    ("手册", "操作说明书", "测评报告/配置核查工具操作说明书.md",
     "测评报告/配置核查工具操作说明书.docx"),
]


def md5(path, n=12):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()[:n]


def docx_stats(path):
    """只读统计：从 docx 包内取页数缓存 / 目录条目 / 内联图 / 表格数。"""
    if not os.path.exists(path):
        return None
    with zipfile.ZipFile(path) as z:
        doc = z.read("word/document.xml").decode("utf-8", "ignore")
        pages = "?"
        if "docProps/app.xml" in z.namelist():
            app = z.read("docProps/app.xml").decode("utf-8", "ignore")
            m = re.search(r"<Pages>(\d+)</Pages>", app)
            if m:
                pages = m.group(1)
            m2 = re.search(r"<Application>([^<]*)</Application>", app)
            app_name = m2.group(1) if m2 else "?"
        else:
            app_name = "?"
        toc = len(re.findall(r'<w:pStyle w:val="TOC\d"', doc))
        figs = len(re.findall(r"<w:drawing>", doc))
        tables = len(re.findall(r"<w:tbl>", doc))
        toc_field = bool(re.search(r"<w:instrText[^>]*>\s*TOC", doc))
    return {
        "pages": pages, "app": app_name, "toc": toc,
        "figs": figs, "tables": tables, "toc_field": toc_field,
    }


def show_stats(path, label, display):
    s = docx_stats(path)
    if not s:
        print(f"  {label} {display}: [缺失] {path}")
        return
    flag = "" if s["toc_field"] else "  [!] 无 TOC 域"
    print(f"  {label} {display}: {s['pages']} 页 | 目录 {s['toc']} 条 | "
          f"图 {s['figs']} | 表 {s['tables']} | {s['app']} | md5 {md5(path)}{flag}")


def build_docx(md_rel, docx_abs):
    """调 md2docx.py 生成 docx。用当前解释器，避免环境错配。"""
    md_abs = os.path.join(ROOT, md_rel)
    if not os.path.exists(md_abs):
        raise FileNotFoundError(md_abs)
    os.makedirs(os.path.dirname(docx_abs), exist_ok=True)
    cmd = [sys.executable, os.path.join(ROOT, "md2docx.py"), md_abs, docx_abs]
    r = subprocess.run(cmd, cwd=ROOT, capture_output=True, text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise RuntimeError(f"md2docx 失败（{r.returncode}）\n{r.stdout[-2000:]}\n{r.stderr[-2000:]}")
    if not os.path.exists(docx_abs):
        raise RuntimeError("md2docx 退出码 0 但没产出文件")


def word_update(docx_abs, pdf=False):
    """Word 打开 -> 更新 TOC 与所有域 -> 保存（可选导出 PDF）-> 读页数。"""
    import win32com.client as win32

    docx_abs = os.path.abspath(docx_abs)
    word = win32.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    doc = None
    try:
        doc = word.Documents.Open(docx_abs, ConfirmConversions=False,
                                  ReadOnly=False, AddToRecentFiles=False)
        # 1) 正文域（PAGEREF/SEQ 等） 2) 目录域缓存
        try:
            doc.Fields.Update()
        except Exception as e:
            print(f"    [warn] Fields.Update: {e}")
        try:
            if doc.TablesOfContents.Count > 0:
                for i in range(1, doc.TablesOfContents.Count + 1):
                    doc.TablesOfContents(i).Update()
        except Exception as e:
            print(f"    [warn] TOC.Update: {e}")
        doc.Save()
        pages = doc.ComputeStatistics(2)  # wdStatisticPages
        toc = 0
        try:
            if doc.TablesOfContents.Count > 0:
                toc = doc.TablesOfContents(1).Range.Paragraphs.Count - 1
        except Exception:
            toc = -1
        if pdf:
            pdf_path = os.path.splitext(docx_abs)[0] + ".pdf"
            doc.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)
            print(f"    PDF: {pdf_path}")
        return pages, toc
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                pass
        try:
            word.Quit()
        except Exception:
            pass


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only-fields", action="store_true", help="不重建，只更新域")
    ap.add_argument("--check", action="store_true", help="只读统计，不开 Word")
    ap.add_argument("--pdf", action="store_true", help="另存 PDF")
    ap.add_argument("--only", choices=["报告", "手册"], help="只处理一份")
    ap.add_argument("--out-dir", help="输出目录（默认测评报告/；验证时可指向临时目录）")
    args = ap.parse_args()

    targets = [r for r in REPORTS if not args.only or r[0] == args.only]
    out_dir = os.path.join(ROOT, args.out_dir) if args.out_dir else None

    if args.check:
        print("=== 只读统计（页数为上次 Word 保存时的缓存值）===")
        for key, disp, _md, docx_rel in targets:
            show_stats(os.path.join(ROOT, docx_rel), key, disp)
            snap = os.path.join(ROOT, "checklist_tool", os.path.basename(docx_rel))
            if os.path.exists(snap):
                same = md5(os.path.join(ROOT, docx_rel)) == md5(snap)
                print(f"      快照副本: {'一致' if same else '[!] 不一致'} "
                      f"({md5(snap)})")
        return 0

    rc = 0
    for key, disp, md_rel, docx_rel in targets:
        src = os.path.join(ROOT, docx_rel)
        dst = os.path.join(out_dir, os.path.basename(docx_rel)) if out_dir else src
        print(f"=== {disp} ===")
        # before 必须在 build 之前取：build 直接覆盖 src，之后再读只会读到新生成的 Pages=1
        before = docx_stats(src)["pages"] if os.path.exists(src) else "?"
        if not args.only_fields:
            print("  [1/2] md -> docx ...")
            build_docx(md_rel, dst)
        elif not os.path.exists(dst):
            print(f"  [skip] 没有现成 docx：{dst}")
            rc = 1
            continue
        print("  [2/2] Word 更新域 ...")
        try:
            pages, toc = word_update(dst, pdf=args.pdf)
        except Exception as e:
            print(f"  [FAIL] Word 更新失败：{e}")
            rc = 1
            continue
        print(f"  页数 {before} -> {pages} | 目录 {toc} 条 | md5 {md5(dst)}")
        if str(before) != str(pages) and before != "?":
            print("  [!] 页数有变化——确认这是预期的（改内容会变页数；"
                  "没改内容却变了，检查 md 是否与交付件不同步）")
    return rc


if __name__ == "__main__":
    sys.exit(main())
