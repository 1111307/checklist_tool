# -*- coding: utf-8 -*-
r"""操作说明书补充配图（第二轮，5 张）。

  op_pack_root.png     3.1 部署：交付包根目录内容（真实 dir）
  op_netdev_list.png   3.5 采集：生成的华为/华三采集命令清单（真实 type）
  op_netdev_dir.png    3.5 采集：回显文件就位后的 netdev 目录（真实 dir）
  op_kylin_report.png  4.4 查看报告：麒麟 HTML 报告在浏览器中的效果
  op_summary_xls.png   六、结果汇总：汇总报告 Excel 打开视图（真实窗口）
"""
import os
import subprocess
import sys

from PIL import Image, ImageChops

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "测评报告", "说明书截图")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
NETDIR = os.path.join(os.path.dirname(HERE), "_netdev_tmp")
XLS = os.path.join(HERE, "win", "output", "配置核查汇总报告_20260911_171023.xls")

sys.path.insert(0, HERE)


def cmd_capture(title, workdir, command, outname, wait=3):
    """真实 cmd 窗口执行命令并截图（复用 _gen_manual_doc_figs 的方式）。"""
    from _gen_manual_doc_figs import capture
    capture(title, workdir, command, outname, wait)


def chrome_shot(url_or_html, outname, w=1200, h=800, scale=2):
    png = os.path.join(OUT, outname)
    if url_or_html.startswith("<"):
        html_path = os.path.join(OUT, "_t.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(url_or_html)
        url = "file:///" + html_path.replace("\\", "/")
    else:
        url = url_or_html
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
                    "--hide-scrollbars", "--force-device-scale-factor=%d" % scale,
                    "--virtual-time-budget=15000", "--run-all-compositor-stages-before-draw",
                    "--window-size=%d,%d" % (w, h), "--screenshot=" + png, url],
                   capture_output=True, text=True)
    if url_or_html.startswith("<") and os.path.exists(os.path.join(OUT, "_t.html")):
        os.remove(os.path.join(OUT, "_t.html"))
    return png


def trim(path, pad=6, white=(255, 255, 255)):
    im = Image.open(path).convert("RGB")
    bg = Image.new("RGB", im.size, white)
    bbox = ImageChops.difference(im, bg).getbbox()
    if bbox:
        im.crop((max(bbox[0] - pad, 0), max(bbox[1] - pad, 0),
                 min(bbox[2] + pad, im.width), min(bbox[3] + pad, im.height))).save(path)
    return Image.open(path).size


def main():
    os.makedirs(OUT, exist_ok=True)

    # ---- 1) 交付包根目录（部署落位）----
    pkg = r"C:\Users\ryan.xiong\Desktop\checklist_tool"
    cmd_capture("说明-交付包目录", pkg, "dir /b", "op_pack_root.png", 3)

    # ---- 2) 网络设备：init 生成清单 + 回显目录 ----
    os.makedirs(os.path.join(NETDIR, "output", "netdev"), exist_ok=True)
    for f in ("run.bat", "check_network.bat", "check_network.ps1"):
        src = os.path.join(HERE, "网络设备核查", f)
        if os.path.exists(src):
            import shutil
            shutil.copy(src, os.path.join(NETDIR, f))
    subprocess.run(["cmd", "/c", "cd /d %s && powershell -NoProfile -ExecutionPolicy Bypass "
                    "-File check_network.ps1 init" % NETDIR],
                   capture_output=True, text=True)
    cmd_capture("说明-采集命令清单", NETDIR,
                'type "output\\netdev\\采集命令清单-华为华三.txt"',
                "op_netdev_list.png", 3)
    # 放入两台设备的回显（真实样例），再截回显目录
    for f in ("dev1_huawei.txt", "dev2_ruijie.txt"):
        src = os.path.join(os.path.dirname(HERE), "_docker_verify", f)
        if os.path.exists(src):
            import shutil
            shutil.copy(src, os.path.join(NETDIR, "output", "netdev",
                                          "CoreSW01.txt" if "huawei" in f else "AggSW02.txt"))
    cmd_capture("说明-回显文件就位", NETDIR,
                "dir /b output\\netdev", "op_netdev_dir.png", 3)

    # ---- 3) 麒麟 HTML 报告在浏览器 ----
    rpt = os.path.join(HERE, "kylin", "output", "配置核查报告_麒麟靶机_20260911_111748.html")
    png = chrome_shot("file:///" + rpt.replace("\\", "/"), "op_kylin_report.png", 1400, 860)
    trim(png, pad=0)

    # ---- 4) 汇总报告 Excel 视图 ----
    if os.path.exists(XLS):
        subprocess.run([sys.executable, os.path.join(HERE, "_capture_excel.py"), XLS,
                        os.path.join(OUT, "op_summary_xls.png"), "24"],
                       capture_output=True, text=True)
        print("汇总报告截图:", os.path.exists(os.path.join(OUT, "op_summary_xls.png")))
    else:
        print("!! 未找到汇总报告:", XLS)

    print("完成")


if __name__ == "__main__":
    main()
