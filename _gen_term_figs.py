# -*- coding: utf-8 -*-
"""终端证据图生成：在真实的 cmd 窗口中执行命令，跑完后按窗口截图，再裁掉下方空白。

产出 9 张图，直接覆盖 测评报告/验证截图/ 下的同名文件：
  term_win_os.png      Windows 版跑操作系统核查（139 项）
  term_win_net.png     Windows 版网络设备核查（init + check）
  term_kylin_os.png    麒麟靶机跑操作系统核查
  probe_win_1_4/1_5/2_11.png   核查项判定 vs 实测配置（Windows）
  probe_kylin_1_1/1_6.png、probe_mysql_1_7.png   同上（麒麟靶机）

要点：
  * 图是**真实窗口截图**（PrintWindow + PW_RENDERFULLCONTENT），不是 HTML 模拟，
    所以能看到真实的系统标题栏、Consolas 字体渲染和滚动后的提示符行；
  * 命令是照实执行的，图里就是真实回显，没有任何合成；
  * 麒麟侧通过 `docker exec` 调用 kylin-target 容器（容器内先手工拉起 mysqld 才有 1.7 的数据）。

依赖：pywin32、Pillow；麒麟侧需 Docker Desktop 与 kylin-target 容器可用。
耗时：Windows 侧因 3 张 probe 图各跑一次完整脚本（约 2 分钟），合计约 11 分钟。

用法：
  python _gen_term_figs.py            # 全部（win + kylin）
  python _gen_term_figs.py win
  python _gen_term_figs.py kylin
"""
import os
import subprocess
import sys
import time

import win32gui
import win32ui
from ctypes import windll
from PIL import Image

BASE = os.path.dirname(os.path.abspath(__file__))
WIN_DIR = os.path.join(BASE, "win")
NETDEV_TEST = os.path.join(BASE, "_netdev_win_test")
OUT = os.path.join(BASE, "测评报告", "验证截图")

WINDOW_SIZE = (1250, 970)      # 截图窗口尺寸（含标题栏）
COLS, LINES = 150, 50          # 控制台字符网格


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
    bits = bmp.GetBitmapBits(True)
    im = Image.frombuffer('RGB', (info['bmWidth'], info['bmHeight']), bits, 'raw', 'BGRX', 0, 1)
    win32gui.DeleteObject(bmp.GetHandle())
    save.DeleteDC()
    mfc.DeleteDC()
    win32gui.ReleaseDC(h, dc)
    return im


def trim_bottom(im, thresh=60, top=34, bottom_margin=14):
    """裁掉下方空白；窗口边框是纯黑、内容背景 #0C0C0C，故阈值取 60 并跳过左右边框。"""
    px = im.load()
    w, h = im.size
    bg = px[w // 2, h - 60]
    for y in range(h - 1 - bottom_margin, top, -1):
        for x in range(16, w - 32, 4):
            p = px[x, y]
            if abs(p[0] - bg[0]) + abs(p[1] - bg[1]) + abs(p[2] - bg[2]) > thresh:
                return im.crop((0, 0, w, min(h, y + 26)))
    return im


def capture(title, workdir, command, outname, wait):
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
    win32gui.MoveWindow(h, 30, 20, WINDOW_SIZE[0], WINDOW_SIZE[1], True)
    time.sleep(2)
    time.sleep(wait)
    out = os.path.join(OUT, outname)
    trim_bottom(grab(h)).save(out)
    print("   saved %s  %s" % (outname, Image.open(out).size))
    win32gui.PostMessage(h, 0x0010, 0, 0)
    time.sleep(1)
    return 0


# ---------------- 任务表 ----------------
CP = "chcp 65001 >nul && "   # 麒麟输出是 UTF-8，先切代码页

JOBS_WIN = [
    ("配置核查工具 - Windows版", WIN_DIR,
     'cscript //Nologo check_xp7.vbs', "term_win_os.png", 150),
    ("网络设备核查（采集-解析）", NETDEV_TEST,
     'powershell -ExecutionPolicy Bypass -File check_network.ps1 init && echo. && '
     'powershell -ExecutionPolicy Bypass -File check_network.ps1 check',
     "term_win_net.png", 40),
    ("核查项 1.4 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[1.4] && echo. && netsh advfirewall show allprofiles',
     "probe_win_1_4.png", 155),
    ("核查项 1.5 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[1.5] && echo. && '
     'reg query HKLM\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters\\Interfaces /s',
     "probe_win_1_5.png", 155),
    ("核查项 2.11 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[2.11] && echo. && net accounts',
     "probe_win_2_11.png", 155),
    ("核查项 1.2 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[1.2] && echo. && '
     'powershell -Command Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct',
     "probe_win_1_2.png", 155),
    ("核查项 2.10 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[2.10] && echo. && '
     'reg query HKLM\\SYSTEM\\CurrentControlSet\\Services\\USBSTOR /v Start',
     "probe_win_2_10.png", 155),
    ("核查项 2.12 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[2.12] && echo. && '
     'sc query WLANSVC && echo. && sc query bthserv',
     "probe_win_2_12.png", 155),
    ("核查项 1.3 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[1.3] && echo. && net start',
     "probe_win_1_3.png", 155),
    ("核查项 1.18 判定与实测", WIN_DIR,
     'cscript //Nologo check_xp7.vbs | findstr /C:[1.18] && echo. && fsutil quota query C:',
     "probe_win_1_18.png", 155),
    # ---- 各组件脚本运行画面（覆盖全部脚本）----
    ("组件核查 - SQL Server", WIN_DIR,
     'set MSSQL_PASS=Check@12345 && cscript //Nologo check_sqlserver.vbs',
     "run_win_sqlserver.png", 45),
    ("组件核查 - MySQL", WIN_DIR,
     'cscript //Nologo check_mysql.vbs', "run_win_mysql.png", 35),
    ("组件核查 - Redis", WIN_DIR,
     'cscript //Nologo check_redis.vbs', "run_win_redis.png", 35),
    ("组件核查 - 达梦", WIN_DIR,
     'cscript //Nologo check_dm.vbs', "run_win_dm.png", 35),
    ("组件核查 - Nginx", WIN_DIR,
     'cscript //Nologo check_nginx.vbs', "run_win_nginx.png", 35),
    ("组件核查 - Tomcat", WIN_DIR,
     'cscript //Nologo check_tomcat.vbs', "run_win_tomcat.png", 35),
]

JOBS_KYLIN = [
    ("麒麟版核查工具 - 银河麒麟靶机", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_kylin.sh", "term_kylin_os.png", 120),
    ("核查项 1.1 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target yum check-update", "probe_kylin_1_1.png", 60),
    ("核查项 1.6 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target grep -i ciphers /etc/ssh/sshd_config",
     "probe_kylin_1_6.png", 25),
    ("核查项 1.7 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target bash /opt/q17.sh", "probe_mysql_1_7.png", 25),
    ("核查项 1.5 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target sestatus", "probe_kylin_1_5.png", 25),
    ("核查项 1.10 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target ss -lntp | findstr 3306", "probe_mysql_1_10.png", 25),
    ("核查项 1.8 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target bash /opt/q18.sh", "probe_mysql_1_8.png", 30),
    ("核查项 1.9 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target bash /opt/q19.sh", "probe_mysql_1_9.png", 30),
    ("核查项 1.12 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target bash /opt/q112.sh", "probe_mysql_1_12.png", 30),
    ("核查项 1.7 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target redis-cli CONFIG GET requirepass", "probe_redis_1_7.png", 25),
    ("核查项 1.10 判定与实测", WIN_DIR,
     CP + "docker exec kylin-target redis-cli CONFIG GET bind", "probe_redis_1_10.png", 25),
    # ---- 各组件脚本运行画面（覆盖全部脚本）----
    ("组件核查 - MySQL", WIN_DIR,
     CP + "docker exec -e MYSQL_USER=checker -e MYSQL_PASS=Checker@123 -w /opt kylin-target bash check_mysql.sh",
     "run_kylin_mysql.png", 45),
    ("组件核查 - Redis", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_redis.sh", "run_kylin_redis.png", 40),
    ("组件核查 - Nginx", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_nginx.sh", "run_kylin_nginx.png", 40),
    ("组件核查 - Tomcat", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_tomcat.sh", "run_kylin_tomcat.png", 40),
    ("组件核查 - 达梦", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_dm.sh", "run_kylin_dm.png", 30),
    ("组件核查 - SQL Server", WIN_DIR,
     CP + "docker exec -w /opt kylin-target bash check_sqlserver.sh", "run_kylin_sqlserver.png", 30),
]


def run(jobs, tag):
    for title, workdir, command, outname, wait in jobs:
        print("=== %s -> %s" % (title, outname), flush=True)
        if workdir and not os.path.isdir(workdir):
            print("   工作目录不存在：%s（网络设备核查需先准备回显目录）" % workdir)
            continue
        rc = capture(title, workdir or BASE, command, outname, wait)
        print("   rc=%d" % rc, flush=True)
        subprocess.call("taskkill /F /IM cmd.exe", shell=True,
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(3)


def main():
    what = sys.argv[1] if len(sys.argv) > 1 else "all"
    os.makedirs(OUT, exist_ok=True)
    if what == "all":
        print("---- Windows 侧（约 11 分钟）----")
        run(JOBS_WIN, "win")
        print("---- 麒麟侧（需 kylin-target 容器与容器内 mysqld）----")
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
