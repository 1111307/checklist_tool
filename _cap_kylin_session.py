# -*- coding: utf-8 -*-
"""麒麟操作画面采集：窗口内只呈现真实的麒麟交互会话。

做法：在麒麟靶机内用 `script` 分配伪终端跑交互式 bash，会话命令由限速喂入器逐条送入，
因此提示符、命令回显与执行结果都是真实顺序。会话第一行是 clear，用于抹掉窗口顶部
启动方式痕迹，窗口里最终只剩会话本身。

用法：
  python _cap_kylin_session.py <输出png> <等待秒数> <命令1> <命令2> ...
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

CONTAINER = "kylin-target"
WINDOW = (1250, 880)
RC = "PS1='[\\u@\\h \\W]# '"
FEED = """#!/bin/bash
while IFS= read -r l; do
  printf '%s\\n' "$l"
  sleep 1.0
done < /tmp/kylin_session.txt
sleep 900
"""
SESSION = """#!/bin/bash
export TERM=xterm
bash /tmp/feed.sh | script -qc 'bash --rcfile /tmp/rc_ps1 -i' /dev/null
"""


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


def main():
    out = os.path.abspath(sys.argv[1])
    wait = int(sys.argv[2])
    lines = sys.argv[3:]

    # 会话脚本与喂入器写进容器
    payload = "\n".join(lines) + "\n"
    subprocess.run(["docker", "exec", "-i", CONTAINER, "bash", "-c",
                    "cat > /tmp/kylin_session.txt"], input=payload.encode("utf-8"), check=False)
    subprocess.run(["docker", "exec", "-i", CONTAINER, "bash", "-c",
                    "cat > /tmp/feed.sh && chmod +x /tmp/feed.sh"], input=FEED.encode(), check=False)
    subprocess.run(["docker", "exec", "-i", CONTAINER, "bash", "-c",
                    "cat > /tmp/session.sh && chmod +x /tmp/session.sh"], input=SESSION.encode(), check=False)
    subprocess.run(["docker", "exec", "-i", CONTAINER, "bash", "-c", "cat > /tmp/rc_ps1"],
                   input=(RC + "\n").encode("utf-8"), check=False)

    cmdline = ('cmd /k "title 麒麟服务器终端 && mode con cols=148 lines=50 && '
               'chcp 65001 >nul && docker exec -i %s bash /tmp/session.sh"' % CONTAINER)
    before = set(list_console())
    subprocess.Popen(cmdline, creationflags=subprocess.CREATE_NEW_CONSOLE)
    time.sleep(6)
    new = [h for h in list_console() if h not in before]
    if not new:
        print("窗口未找到")
        return 2
    h = new[0]
    win32gui.MoveWindow(h, 40, 25, WINDOW[0], WINDOW[1], True)
    win32gui.SetWindowText(h, "麒麟服务器终端")
    time.sleep(wait)

    if not win32gui.IsWindow(h):
        cand = [x for x in list_console() if win32gui.GetWindowText(x).strip() == "麒麟服务器终端"]
        if not cand:
            print("窗口已关闭")
            return 3
        h = cand[0]
    win32gui.ShowWindow(h, win32con.SW_RESTORE)
    win32gui.MoveWindow(h, 40, 25, WINDOW[0], WINDOW[1], True)
    win32gui.SetWindowText(h, "麒麟服务器终端")
    time.sleep(2)
    im = grab(h)
    os.makedirs(os.path.dirname(out), exist_ok=True)
    im.save(out)
    print("saved", os.path.basename(out), im.size)
    win32gui.PostMessage(h, 0x0010, 0, 0)
    # 收尾：把容器里的会话进程停掉
    time.sleep(1)
    subprocess.run(["docker", "exec", CONTAINER, "pkill", "-f", "feed.sh"], check=False)
    subprocess.run(["docker", "exec", CONTAINER, "pkill", "-f", "script -qc"], check=False)
    return 0


if __name__ == "__main__":
    sys.exit(main())
