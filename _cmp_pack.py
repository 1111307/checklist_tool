# -*- coding: utf-8 -*-
"""交付包一致性校验（两段检查，缺一不可）。

A. 仓库根 vs checklist_tool/ 快照：只比**公共成品**。
   交付包不放源 .md，也不放核查表标注验证副本，所以这里不是全目录镜像；
   只约束报告 docx / 指导书 / 核查表 / 脚本 / 人工核查台 / README。
   例外：docx 若一边被 WPS 重存过，字节必然不同但内容相同——退化为逐段文本
   比对，内容一致即视为通过。

B. 仓库 vs **最终交付目录** 交付包_v2.0.0_20261025/：逐文件比字节。
   这是真正要交给用户的目录，以前只做 A 段检查，包里少文件/放了旧版都发现不了。
   《交付说明.docx》是交付包独有产物（源 md 在仓库根），列为「包内独有」不报错。

用法
----
    python _cmp_pack.py            # A + B 全查；退出码 0=全部一致
    python _cmp_pack.py -v         # 逐个文件打印
    python _cmp_pack.py --strict   # A 段字节级严格模式
    python _cmp_pack.py --skip-pack  # 只跑 A 段
"""
import argparse
import hashlib
import os
import re
import sys
import zipfile

try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

ROOT = os.path.dirname(os.path.abspath(__file__))
SNAP = os.path.join(ROOT, "checklist_tool")
PACK = os.path.join(ROOT, "交付包_v2.0.0_20261025")

# 公共成品：根路径 -> 快照内相对路径
PUBLIC = [
    "README.md",
    "配置核查作业指导书_v2.0.0.docx",
    "配置核查表_v2.0.0.xlsx",
    "配置核查表_v2.0.0_标注自动验证.xlsx",
    "manual_check.html",
    "指导书与核查工具交叉验证报告.docx",
    "配置核查工具操作说明书.docx",
]

# 目录：根目录 -> 快照内目录，按扩展名收脚本（不含 output/）
SCRIPT_DIRS = [
    ("win", "win", (".vbs", ".ps1", ".bat", ".conf")),
    ("kylin", "kylin", (".sh", ".conf")),
    ("网络设备核查", "网络设备核查", (".sh", ".ps1", ".bat", ".conf")),
]

# 报告在仓库根不在根目录，单独映射
REPORT_MAP = {
    "指导书与核查工具交叉验证报告.docx": "测评报告/指导书与核查工具交叉验证报告.docx",
    "配置核查工具操作说明书.docx": "测评报告/配置核查工具操作说明书.docx",
}


def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def docx_text_hash(path):
    """docx 正文纯文本哈希 + 段数（用于判定「打包不同但内容相同」）。"""
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml").decode("utf-8", "ignore")
    paras = re.findall(r"<w:p[ >].*?</w:p>", xml, re.S)
    texts = ["".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", p)) for p in paras]
    return hashlib.md5("\n".join(texts).encode("utf-8")).hexdigest()[:16], len(texts)


def collect():
    """返回 [(显示名, 根绝对路径, 快照绝对路径)]。"""
    out = []
    for name in PUBLIC:
        root_path = os.path.join(ROOT, REPORT_MAP.get(name, name))
        out.append((name, root_path, os.path.join(SNAP, name)))
    for root_dir, snap_dir, exts in SCRIPT_DIRS:
        d = os.path.join(ROOT, root_dir)
        if not os.path.isdir(d):
            continue
        for fn in sorted(os.listdir(d)):
            if fn.lower().endswith(exts):
                out.append((f"{root_dir}/{fn}",
                            os.path.join(d, fn),
                            os.path.join(SNAP, snap_dir, fn)))
    return out


def find_repo_counterpart(rel):
    """交付包内相对路径 -> 仓库中对应的源文件路径。"""
    for cand in (os.path.join(ROOT, rel),
                 os.path.join(ROOT, "测评报告", os.path.basename(rel)),
                 os.path.join(SNAP, rel)):
        if os.path.isfile(cand):
            return cand
    return None


def check_pack(verbose=False):
    """B 段：交付目录与仓库逐字节比对。返回 0/1。"""
    if not os.path.isdir(PACK):
        print(f"\n[B] [FAIL] 找不到交付目录：{PACK}")
        return 1
    files = []
    for dp, _dn, fns in os.walk(PACK):
        for fn in fns:
            files.append(os.path.relpath(os.path.join(dp, fn), PACK))
    same, diff, only = [], [], []
    for rel in sorted(files):
        src = find_repo_counterpart(rel)
        dst = os.path.join(PACK, rel)
        if src is None:
            only.append(rel)
            continue
        if md5(src) == md5(dst):
            same.append(rel)
            if verbose:
                print(f"  SAME  {rel}")
        else:
            diff.append((rel, os.path.relpath(src, ROOT), md5(src)[:12], md5(dst)[:12]))
    print(f"\n=== [B] 交付目录一致性：{len(files)} 个文件 ===")
    print(f"  与仓库逐字节一致 {len(same)}")
    if only:
        print(f"  交付包独有（无仓库副本，正常） {len(only)}")
        for rel in only:
            print(f"      {rel}")
    if diff:
        print(f"  [不一致] {len(diff)}")
        for rel, src, ha, hb in diff:
            print(f"      {rel}  <- {src}  {ha} / {hb}")
    if diff:
        print("结论：交付目录与仓库不一致 -> 退出码 1")
        return 1
    print("结论：交付目录与仓库一致 -> 退出码 0")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", "--verbose", action="store_true")
    ap.add_argument("--strict", action="store_true", help="A 段字节级严格模式")
    ap.add_argument("--skip-pack", action="store_true", help="只跑 A 段")
    args = ap.parse_args()

    if not os.path.isdir(SNAP):
        print(f"[FAIL] 找不到快照目录：{SNAP}")
        return 1

    same, diff, missing, soft = [], [], [], []
    for name, a, b in collect():
        if not os.path.exists(a):
            missing.append((name, "仓库根缺文件", a))
            continue
        if not os.path.exists(b):
            missing.append((name, "快照缺文件", b))
            continue
        ha, hb = md5(a), md5(b)
        if ha == hb:
            same.append(name)
            if args.verbose:
                print(f"  SAME  {name}")
            continue
        # 字节不同：docx 再比内容
        if not args.strict and name.lower().endswith(".docx"):
            ta, na = docx_text_hash(a)
            tb, nb = docx_text_hash(b)
            if ta == tb and na == nb:
                soft.append((name, na, ha[:12], hb[:12]))
                if args.verbose:
                    print(f"  SOFT  {name}（打包不同，内容一致，{na} 段）")
                continue
            diff.append((name, f"内容不同（段数 {na} vs {nb}）", ha[:12], hb[:12]))
            continue
        diff.append((name, "字节不同", ha[:12], hb[:12]))

    total = len(same) + len(diff) + len(missing) + len(soft)
    print(f"\n=== 交付包一致性：{total} 项 ===")
    print(f"  一致      {len(same)}")
    if soft:
        print(f"  内容一致  {len(soft)}（字节不同：WPS/Word 打包差异，允许）")
        for name, n, ha, hb in soft:
            print(f"      {name}  {n} 段  {ha} / {hb}")
    if diff:
        print(f"  [不一致]  {len(diff)}")
        for name, why, ha, hb in diff:
            print(f"      {name}  {why}  {ha} / {hb}")
    if missing:
        print(f"  [缺失]    {len(missing)}")
        for name, why, p in missing:
            print(f"      {name}  {why}  {p}")

    rc = 1 if (diff or missing) else 0
    print("\n结论：公共成品" + ("存在不一致" if rc else "全部一致"))
    if not args.skip_pack:
        rc |= check_pack(args.verbose)
    return rc


if __name__ == "__main__":
    sys.exit(main())
