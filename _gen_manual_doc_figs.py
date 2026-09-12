# -*- coding: utf-8 -*-
"""《配置核查工具操作说明书》配图生成（全部为真实运行截图）。

Windows 侧：真实 cmd 窗口执行真实命令后截图。
麒麟侧：在麒麟靶机容器内执行（与报告中麒麟配图同一方式）。

用法：
  python _gen_manual_doc_figs.py            # 全部
  python _gen_manual_doc_figs.py win        # 只 Windows
  python _gen_manual_doc_figs.py kylin      # 只麒麟
  python _gen_manual_doc_figs.py <关键词>    # 只跑名字含关键词的
"""
import os
import subprocess
import sys
import time

import win32gui
import win32ui
import win32con
from ctypes import windll
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
WIN = os.path.join(BASE, "win")
OUT = os.path.join(BASE, "测评报告", "说明书截图")
CHROME_UNUSED = None

WINDOW_SIZE = (1250, 940)
COLS, LINES = 150, 52


def list_console():
    found = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == 'ConsoleWindowClass':
            found.append(h)

    win32gui.EnumWindows(cb, None)
    return found


def grab(h):
    l, t, r, b = win32gui.GetWindowRect(h)
    w, hh = r - l, b - t
    dc = win32gui.GetWindowDC(h)
    mfc = win32ui.CreateDCFromHandle(dc)
    save = mfc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfc, w, hh)
    save.SelectObject(bmp)
    windll.user32.PrintWindow(h, save.GetSafeHdc(), 2)
    info = bmp.GetInfo()
    im = Image.frombuffer('RGB', (info['bmWidth'], info['bmHeight']),
                          bmp.GetBitmapBits(True), 'raw', 'BGRX', 0, 1)
    win32gui.DeleteObject(bmp.GetHandle())
    save.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(h, dc)
    return im


def trim_bottom(im, thresh=60, top=34, bottom_margin=14):
    px = im.load()
    w, h = im.size
    bg = px[w // 2, h - 60]
    for y in range(h - 1 - bottom_margin, top, -1):
        for x in range(16, w - 32, 4):
            p = px[x, y]
            if abs(p[0] - bg[0]) + abs(p[1] - bg[1]) + abs(p[2] - bg[2]) > thresh:
                return im.crop((0, 0, w, min(h, y + 26)))
    return im


def capture(title, workdir, command, outname, wait, tail_only=0):
    cmdline = 'cmd /k "title {} && mode con cols={} lines={} && cd /d {} && {}"'.format(
        title, COLS, LINES, workdir, command)
    before = set(list_console())
    subprocess.Popen(cmdline, creationflags=subprocess.CREATE_NEW_CONSOLE)
    time.sleep(6)
    new = [h for h in list_console() if h not in before]
    if not new:
        print("   窗口未找到")
        return 2
    h = new[0]
    win32gui.MoveWindow(h, 40, 25, WINDOW_SIZE[0], WINDOW_SIZE[1], True)
    time.sleep(2)
    time.sleep(wait)
    # 等待期间窗口可能被最小化：截图前恢复并重新定位
    if not win32gui.IsWindow(h):
        cand = [x for x in list_console() if win32gui.GetWindowText(x).strip() == title]
        if not cand:
            print("   窗口已关闭，跳过 %s" % outname)
            return 3
        h = cand[0]
    win32gui.ShowWindow(h, win32con.SW_RESTORE)
    win32gui.MoveWindow(h, 40, 25, WINDOW_SIZE[0], WINDOW_SIZE[1], True)
    time.sleep(2)
    out = os.path.join(OUT, outname)
    trim_bottom(grab(h)).save(out)
    print("   saved %s %s" % (outname, Image.open(out).size))
    win32gui.PostMessage(h, 0x0010, 0, 0)
    time.sleep(1)
    return 0


CP = "chcp 65001 >nul && "
D = "docker exec "
DD = "docker exec -w /opt "

# ---------- Windows 侧 ----------
JOBS_WIN = [
    # 1. 工具包文件清单
    ("说明-工具包文件", WIN, 'dir /b *.bat *.vbs *.ps1 *.conf', "op_win_files.png", 4),
    # 2. 连接参数配置文件
    ("说明-连接参数配置", WIN, 'type db_config.conf', "op_win_conf.png", 4),
    # 3. 一键运行（真跑，7 个脚本 + 网络设备 + 汇总）
    ("说明-一键运行", WIN, 'run_all.bat', "op_win_runall.png", 210),
    # 4. 报告输出清单
    ("说明-报告输出", WIN, 'dir /o-d /b output\\*.html output\\*.xls', "op_win_output.png", 4),
]

# ---------- 麒麟侧 ----------
JOBS_KYLIN = [
    # 1. 自动探测与执行过程（真跑）
    ("说明-麒麟一键运行", WIN,
     CP + DD + "kylin-target bash run_all.sh", "op_kylin_runall.png", 240),
    # 2. 报告输出清单
    ("说明-麒麟报告输出", WIN,
     CP + D + "kylin-target ls -lht /opt/output", "op_kylin_output.png", 20),
]


def run(jobs, tag):
    for title, workdir, command, outname, wait in jobs:
        print("=== %s -> %s" % (title, outname), flush=True)
        rc = capture(title, workdir, command, outname, wait)
        print("   rc=%d" % rc, flush=True)
        subprocess.call("taskkill /F /IM cmd.exe", shell=True,
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(3)


def main():
    what = sys.argv[1] if len(sys.argv) > 1 else "all"
    os.makedirs(OUT, exist_ok=True)
    if what == "all":
        run(JOBS_WIN, "win")
        run(JOBS_KYLIN, "kylin")
    elif what == "win":
        run(JOBS_WIN, "win")
    elif what == "kylin":
        run(JOBS_KYLIN, "kylin")
    else:
        sel = [j for j in (JOBS_WIN + JOBS_KYLIN) if what in j[3]]
        print("按关键词 %s 选中 %d 项" % (what, len(sel)))
        run(sel, what)
    print("完成")


if __name__ == "__main__":
    main()
