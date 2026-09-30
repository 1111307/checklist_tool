# -*- coding: utf-8 -*-
"""评审整改第 4/8/9 条的数据源：从指导书 md 提取「每条核查项的验证过程/方法」。

输出 _method_map.json：
  { "1.4": {"title": "操作系统应具备防火墙功能",
            "method": "1.4.1 Windows：netsh advfirewall show allprofiles；1.4.2 麒麟：systemctl status firewalld、firewall-cmd --state"},
    ... }

规则：
  - H2（## N.M 标题）= 核查项；其下直到下一个 H2 的 H3（### N.M.k 平台）及其正文 = 方法。
  - 方法正文里被「中文引号」包住的命令串就是指导书给定的核查命令，逐个抽出；
    连同 H3 小节号与平台名拼成一行，多平台用「；」连接。
  - 没有 H3 小节的章（7/8/9/10 章组织管理类、第 5 章网络设备类）给固定人工口径。
  - 命令串里去掉明显的叙述词，只留命令本体（含参数）。
"""
import json
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

SRC = "配置核查作业指导书_v2.0.0.md"
OUT = "_method_map.json"

MANUAL_DEFAULT = "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）"
NETDEV_DEFAULT = ("采集-解析：按采集清单登录设备执行只读命令、保存回显，"
                  "脚本解析回显判定（命令清单见 check_network 采集模板）")

text = open(SRC, encoding="utf-8").read()

# 按 H2 切块
h2_pat = re.compile(r"^## (\d+\.\d+) (.+)$", re.M)
matches = list(h2_pat.finditer(text))
items = {}
for i, m in enumerate(matches):
    cid, title = m.group(1), m.group(2).strip()
    body = text[m.end(): matches[i + 1].start() if i + 1 < len(matches) else len(text)]
    # 该块内的 H3 小节
    h3_pat = re.compile(r"^### (\S+)\s*(.*)$", re.M)
    h3s = list(h3_pat.finditer(body))
    if not h3s:
        chap = cid.split(".")[0]
        items[cid] = {
            "title": title,
            "method": NETDEV_DEFAULT if chap == "5" else MANUAL_DEFAULT,
            "src": "no-h3",
        }
        continue
    parts = []
    for j, h3 in enumerate(h3s):
        sec_no = h3.group(1)
        plat = re.sub(r"[WindowsXP7麒麟中标银河、/\s]+", "", h3.group(2)) or h3.group(2).strip()
        plat_raw = h3.group(2).strip()
        seg = body[h3.end(): h3s[j + 1].start() if j + 1 < len(h3s) else len(body)]
        # 正文直到下一个 H3 为止；H4（#### ）也算边界
        h4 = re.search(r"^#### ", seg, re.M)
        if h4:
            seg = seg[: h4.start()]
        cmds = re.findall(r"“([^”]{2,120})”", seg)
        # 过滤：只留像命令的（含拉丁/斜杠且不是纯中文叙述）
        keep = []
        for c in cmds:
            c = c.strip()
            if not re.search(r"[A-Za-z$/-]", c):
                continue
            if re.match(r"^(控制面板|桌面|开始|我的电脑|任务栏|运行)", c):
                continue
            if c not in keep:
                keep.append(c)
        cmd_str = "、".join(keep[:6])
        if cmd_str:
            parts.append("%s %s：%s" % (sec_no, plat_raw, cmd_str))
        else:
            parts.append("%s %s：图形界面/文档核查" % (sec_no, plat_raw))
    items[cid] = {"title": title, "method": "；".join(parts), "src": "guide"}

# 标题对齐用：去掉句读差异（报告里的核查项名一律用指导书 H2 原文）
with open(OUT, "w", encoding="utf-8") as f:
    json.dump(items, f, ensure_ascii=False, indent=1)

n_h3 = sum(1 for v in items.values() if v["src"] == "guide")
n_man = sum(1 for v in items.values() if v["method"] == MANUAL_DEFAULT)
n_net = sum(1 for v in items.values() if v["method"] == NETDEV_DEFAULT)
print("提取 %d 条：指导书方法 %d、人工口径 %d、网络设备口径 %d" % (len(items), n_h3, n_man, n_net))
from collections import Counter
print("分章：", dict(sorted(Counter(k.split(".")[0] for k in items).items())))
# 抽查 1.4
print("\n1.4 示例：", items.get("1.4", {}).get("method", "")[:200])
print("\n2.6 示例：", items.get("2.6", {}).get("method", "")[:160])
