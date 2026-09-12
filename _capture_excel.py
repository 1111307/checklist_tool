# -*- coding: utf-8 -*-
"""打开人工核查台导出的 .xls，截取 Excel 窗口并裁到数据区（真实截图）。

用法：python _capture_excel.py <xls路径> <输出png> [裁剪到的行号]
"""
import os
import sys
import time

import win32gui
import win32ui
import win32con
from ctypes import windll
from PIL import Image
import win32com.client as win32

# 让进程 DPI 感知，否则 PointsToScreenPixels 与截图像素对不上
try:
    windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def find_window(cls, timeout=45):
    found = []

    def cb(h, _):
        if win32gui.IsWindowVisible(h) and win32gui.GetClassName(h) == cls:
            found.append(h)

    end = time.time() + timeout
    while time.time() < end:
        found.clear()
        win32gui.EnumWindows(cb, None)
        if found:
            return found[0]
        time.sleep(1)
    return 0


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
    return im, (l, t)


def main():
    xls, out = sys.argv[1], sys.argv[2]
    last_row = int(sys.argv[3]) if len(sys.argv) > 3 else 26
    xls = os.path.abspath(xls)

    app = win32.DispatchEx("Excel.Application")
    app.Visible = True
    app.DisplayAlerts = False
    wb = app.Workbooks.Open(xls)
    ws = wb.Worksheets(1)
    ws.Columns("D").ColumnWidth = 40      # 核查项
    ws.Columns("F").ColumnWidth = 34      # 详情
    ws.Columns("G").ColumnWidth = 24      # 建议
    ws.Columns("H").ColumnWidth = 26      # 参考指导书
    ws.Columns("I:Z").Hidden = True       # 表格只用 A~H，隐藏其余列以免白白占宽
    ws.Range("A1").Select()

    h = find_window("XLMAIN")
    if not h:
        print("未找到 Excel 窗口")
        return 2
    # 还原后调成贴合表格的窗口大小：图幅越窄，插入文档后表内文字越大
    win32gui.ShowWindow(h, win32con.SW_RESTORE)
    time.sleep(1)
    win32gui.MoveWindow(h, 30, 20, 1340, 860, True)
    time.sleep(2)

    im, origin = grab(h)

    # 裁掉表格下方的空白区：自下而上找最后一行含非白像素的行。
    # 底部还有横向滚动条与状态栏（都算非内容），从它们上方起扫。
    px = im.load()
    w, hh = im.size
    bottom = hh - 1
    for y in range(hh - 105, 100, -1):
        row_hit = False
        for x in range(40, w - 40, 6):
            r_, g_, b_ = px[x, y][:3]
            if r_ < 180 or g_ < 180 or b_ < 180:   # 网格线约 217，不算内容
                row_hit = True
                break
        if row_hit:
            bottom = y
            break
    im = im.crop((0, 0, w, min(hh, bottom + 30)))

    im.save(out)
    print("saved", out, im.size)

    wb.Close(False)
    app.Quit()
    time.sleep(1)
    return 0


if __name__ == "__main__":
    sys.exit(main())
