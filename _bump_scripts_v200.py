# -*- coding: utf-8 -*-
r"""把核查脚本里的指导书版本字样 v2.2 → v2.0.0（字节级替换，保持原编码与换行）。

覆盖：仓库根 win/ kylin/ 网络设备核查/ + checklist_tool/ 快照 + Desktop 实物包对应目录。
只替换 ASCII 序列 b"v2.2"，不动其它任何字节。
"""
import glob
import hashlib
import os

ROOTS = [
    "",
    "checklist_tool/",
    r"C:/Users/ryan.xiong/Desktop/checklist_tool/",
]
PATTERNS = ["win/*.vbs", "win/*.ps1", "win/*.bat", "kylin/*.sh",
            "网络设备核查/*.ps1", "网络设备核查/*.sh", "网络设备核查/*.bat", "网络设备核查/*.md"]

OLD, NEW = b"v2.2", b"v2.0.0"
changed = []

for root in ROOTS:
    for pat in PATTERNS:
        for f in sorted(glob.glob(os.path.join(root, pat))):
            raw = open(f, "rb").read()
            n = raw.count(OLD)
            if not n:
                continue
            open(f, "wb").write(raw.replace(OLD, NEW))
            changed.append((f, n))
            print("  %-52s %d 处" % (f.replace("C:/Users/ryan.xiong/Desktop/checklist_tool/", "Desktop/"), n))

print("\n共处理 %d 个文件，%d 处替换" % (len(changed), sum(n for _, n in changed)))

# 同源副本一致性（仓库根 vs 快照 vs 实物包）
print("\n同源副本字节一致性：")
pairs = [
    ("win/check_xp7.vbs", "checklist_tool/win/check_xp7.vbs"),
    ("kylin/check_mysql.sh", "checklist_tool/kylin/check_mysql.sh"),
    ("网络设备核查/check_network.ps1", "checklist_tool/网络设备核查/check_network.ps1"),
    ("win/check_xp7.vbs", r"C:/Users/ryan.xiong/Desktop/checklist_tool/win/check_xp7.vbs"),
    ("kylin/check_mysql.sh", r"C:/Users/ryan.xiong/Desktop/checklist_tool/kylin/check_mysql.sh"),
]
for a, b in pairs:
    if not (os.path.exists(a) and os.path.exists(b)):
        print("  跳过（缺文件）:", a, "|", b)
        continue
    ha = hashlib.md5(open(a, "rb").read()).hexdigest()[:10]
    hb = hashlib.md5(open(b, "rb").read()).hexdigest()[:10]
    print("  %-40s %s %s %s" % (os.path.basename(a) + " vs " + b.split("/")[-2] + "/", ha, hb, "一致" if ha == hb else "!! 不一致"))

# 残留检查
print("\n残留 v2.2：")
left = []
for root in ROOTS:
    for pat in PATTERNS:
        for f in sorted(glob.glob(os.path.join(root, pat))):
            if b"v2.2" in open(f, "rb").read():
                left.append(f)
print("  " + ("无" if not left else "\n  ".join(left)))
