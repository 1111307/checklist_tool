# -*- coding: utf-8 -*-
r"""去掉麒麟 HTML 报告截图头部重复的「参考标准」行（图内修补，不改脚本）。

原理：头部蓝色区文字为白字，按行统计亮像素数高于基线即为文字行；
头部共 4 行文字（标题、参考标准副标题、系统信息、重复的参考标准），
对最后一行所在行带做逐列线性插值（用其上方与下方各一条净空行），
使渐变无缝，从而抹掉重复行。

用法：python _fix_kylin_fig.py [--apply]
"""
import sys

import numpy as np
from PIL import Image

FIGS = [
    "测评报告/说明书截图/op_kylin_report.png",
    "测评报告/验证截图/fig_kylin_os.png",
    "测评报告/验证截图/fig_kylin_mysql.png",
    "测评报告/验证截图/fig_kylin_nginx.png",
    "测评报告/验证截图/fig_kylin_redis.png",
    "测评报告/验证截图/fig_kylin_tomcat.png",
    "测评报告/验证截图/fig_kylin_net.png",
]
APPLY = "--apply" in sys.argv


def find_dup_line(a):
    """返回（重复行起止 y, 头部蓝色区起止 y），找不到返回 None。"""
    h, w, _ = a.shape
    bright = ((a > 170).all(axis=2)).sum(axis=1)
    blue = (a[..., 2].astype(int) - a[..., 0].astype(int) > 30).mean(axis=1)
    hdr_rows = np.where(blue > 0.5)[0]
    if len(hdr_rows) < 50:
        return None
    top, bot = int(hdr_rows[0]), int(hdr_rows[-1])
    seg = bright[top:bot + 1]
    base = int(np.median(seg))
    thr = base * 1.10
    lines, run = [], None
    for i, v in enumerate(seg):
        if v > thr:
            run = i if run is None else run
        elif run is not None:
            if i - run >= 4:                      # 至少 4 行才算一条文字行
                lines.append((top + run, top + i - 1))
            run = None
    if run is not None and len(seg) - run >= 4:
        lines.append((top + run, top + len(seg) - 1))
    if len(lines) < 4:
        return None, (top, bot), len(lines)
    return lines[-1], (top, bot), len(lines)


def patch(path, dry=True):
    im = Image.open(path).convert("RGB")
    a = np.array(im)
    res = find_dup_line(a)
    if res is None:
        print("%-52s 非报告截图（无蓝色头部），跳过" % path)
        return
    line, (top, bot), n = res
    if line is None:
        print("%-52s 头部文字行仅 %d 条，未见重复行，跳过" % (path, n))
        return
    y0, y1 = line
    pad_up, pad_dn = 3, 5
    ys, ye = max(top + 1, y0 - pad_up), min(bot - 1, y1 + pad_dn)
    up_row, dn_row = ys - 2, ye + 2
    if up_row <= top or dn_row >= bot:
        print("%-52s 上下净空不足，跳过" % path)
        return
    print("%-52s 重复行 y=%d..%d（头部 %d..%d，共 %d 行文字）→ 用 y=%d/%d 插值填补装 %d..%d"
          % (path, y0, y1, top, bot, n, up_row, dn_row, ys, ye))
    if dry:
        return
    up = a[up_row].astype(float)
    dn = a[dn_row].astype(float)
    span = dn_row - up_row
    for y in range(ys, ye + 1):
        t = (y - up_row) / span
        a[y] = np.clip(up * (1 - t) + dn * t, 0, 255).astype(np.uint8)
    Image.fromarray(a).save(path)
    print("    已修补")


for f in FIGS:
    try:
        patch(f, dry=not APPLY)
    except FileNotFoundError:
        print("%-52s 不存在，跳过" % f)
if not APPLY:
    print("（干跑；加 --apply 落盘）")
