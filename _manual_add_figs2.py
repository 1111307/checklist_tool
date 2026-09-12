# -*- coding: utf-8 -*-
r"""操作说明书加图（第二轮）：插入 5 张新图并重排图号（15→20）。

新图：
  图 2   交付包根目录内容          op_pack_root.png    （4.1 部署）
  图 6   生成的采集命令清单        op_netdev_list.png  （4.5 网络设备核查）
  图 7   回显文件就位后的目录      op_netdev_dir.png   （4.5）
  图 16  麒麟侧 HTML 报告样例      op_kylin_report.png （5.4 查看报告）
  图 20  汇总报告 Excel 视图       op_summary_xls.png  （八、结果汇总与上报）
"""
import pathlib
import re

MD = pathlib.Path("测评报告/配置核查工具操作说明书.md")
s = MD.read_text(encoding="utf-8")

# ---------- 1) 既有图号重排 ----------
MAPPING = {1: 1, 2: 3, 3: 4, 4: 5, 5: 8, 6: 9, 7: 10, 8: 11,
           9: 12, 10: 13, 11: 14, 12: 16, 13: 17, 14: 18, 15: 19}


def renum(m):
    return "![图 %d　" % MAPPING.get(int(m.group(1)), int(m.group(1)))


s = re.sub(r"!\[图 (\d+)　", renum, s)


def block(n, cap, fig, note):
    return "![图 %d　%s](说明书截图/%s)\n\n> %s\n\n" % (n, cap, fig, note)


def insert_after(anchor, payload, label):
    global s
    assert anchor in s, "锚点未找到: " + label
    s = s.replace(anchor, anchor + "\n" + payload, 1)
    print("已插入", label)


def insert_before(anchor, payload, label):
    global s
    assert anchor in s, "锚点未找到: " + label
    s = s.replace(anchor, payload + anchor, 1)
    print("已插入", label)


# ---------- 2) 插入新图 ----------
insert_after(
    "可先将该目录加入白名单，或改用管理员命令行运行。",
    block(2, "交付包根目录内容（部署前核对）", "op_pack_root.png",
          "在交付包根目录执行 dir /b 的实际输出。拷贝前先核对包内容完整：win 与 kylin 两个脚本目录、"
          "网络设备核查独立版、人工核查台 manual_check.html，以及指导书、核查表、操作说明书与交叉验证报告等文档。"),
    "图 2 交付包根目录")

insert_before(
    "![图 8　网络设备核查的采集与解析](说明书截图/term_win_net.png)",
    block(6, "生成的采集命令清单（华为/华三模板）", "op_netdev_list.png",
          "执行 init 后 output\\netdev 下生成的采集命令清单实际内容，逐行列出需登录设备执行的命令"
          "（版本、配置、VLAN、路由、SSH 状态、账户清单、ACL、802.1x、端口隔离、组播表等），"
          "防火墙设备另有补充命令。")
    + "\n"
    + block(7, "回显文件就位后的 netdev 目录", "op_netdev_dir.png",
            "运维登录设备执行清单命令后，将回显按设备名保存为 output\\netdev\\<设备名>.txt"
            "（图中 CoreSW01.txt、AggSW02.txt 为两台交换机的回显），即可运行 check 解析出第 5 章 23 项报告。"),
    "图 6/7 网络设备采集")

insert_after(
    "![图 14　麒麟侧报告输出](说明书截图/op_kylin_output.png)",
    "\n" + block(15, "麒麟侧 HTML 报告样例", "op_kylin_report.png",
                 "麒麟侧生成的 HTML 报告在浏览器中打开的实际效果：页首为系统信息与四态统计卡，"
                 "下方为逐项明细表，含判定、检测详情、整改建议与对应指导书条款。"),
    "图 15 麒麟 HTML 报告")

insert_after(
    "逐项可追溯。",
    "\n" + block(20, "汇总报告 Excel 视图（上报用）", "op_summary_xls.png",
                 "Windows 版一键运行合并生成的汇总报告在 Excel 中打开的实际效果：表头标注组件数，"
                 "表体按组件列出全部检查项的判定、详情、修复建议与对应指导书章节，日常上报以此文件为准。"),
    "图 20 汇总报告视图")

MD.write_text(s, encoding="utf-8")

figs = [int(m.group(1)) for m in re.finditer(r"!\[图 (\d+)　", s)]
print("\n图号序列:", figs)
print("连续 1-20:", figs == list(range(1, 21)))
