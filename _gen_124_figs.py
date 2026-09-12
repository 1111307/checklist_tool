# -*- coding: utf-8 -*-
r"""按指导书现有风格（白底黑字等宽终端输出）生成 3 张核查示例图。

产出（_tmp_imgs/）：
  gen_dm_win.png      达梦 Windows：cmd 下登录 disql（对应 1.24.2.3 Windows 方法段）
  gen_redis_win.png   Redis Windows：redis-cli.exe 连接 + 日志配置查看
  gen_redis_kylin.png Redis 麒麟：终端 redis-cli + 慢查询/日志核查
"""
import os
import subprocess

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "_tmp_imgs")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CSS = ('body{margin:0;background:#fff;}'
       'pre{font:15px/1.9 "Courier New","Consolas",monospace;color:#000;'
       'margin:0;padding:6px 10px;white-space:pre;}')


def render(name, lines):
    body = "\n".join(lines).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    html_path = os.path.join(OUT, name + ".html")
    png_path = os.path.join(OUT, name + ".png")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>%s</style>'
                '</head><body><pre>%s</pre></body></html>' % (CSS, body))
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
                    "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--virtual-time-budget=3000", "--window-size=1200,600",
                    "--screenshot=" + png_path, "file:///" + html_path.replace("\\", "/")],
                   capture_output=True, text=True)
    im = Image.open(png_path).convert("RGB")
    px = im.load()
    w, h = im.size
    last = h - 1
    for y in range(h - 1, 10, -1):
        if any(px[x, y] != (255, 255, 255) for x in range(10, w - 10, 6)):
            last = y
            break
    im.crop((0, 0, w, min(h, last + 14))).save(png_path)
    print("生成", name + ".png", Image.open(png_path).size)
    os.remove(html_path)


render("gen_dm_win", [
    r"C:\Users\Administrator> cd /d C:\dmdbms\bin",
    "",
    r"C:\dmdbms\bin> disql SYSDBA/SYSDBA@localhost:5236",
    "",
    "服务器[localhost:5236]:处于普通打开状态",
    "登录使用时间: 15.678(ms)",
    "disql V8",
    "SQL>",
])

render("gen_redis_win", [
    r"C:\Users\Administrator> cd C:\Redis",
    "",
    r"C:\Redis> redis-cli.exe -a ******",
    "127.0.0.1:6379> CONFIG GET logfile",
    '1) "logfile"',
    r'2) "C:\Redis\redis.log"',
    "127.0.0.1:6379> CONFIG GET loglevel",
    '1) "loglevel"',
    '2) "notice"',
    "127.0.0.1:6379>",
])

render("gen_redis_kylin", [
    "[root@kylin-server ~]# redis-cli -a ******",
    "127.0.0.1:6379> CONFIG GET logfile",
    '1) "logfile"',
    '2) "/var/log/redis/redis.log"',
    "127.0.0.1:6379> CONFIG GET slowlog-log-slower-than",
    '1) "slowlog-log-slower-than"',
    '2) "10000"',
    "127.0.0.1:6379> SLOWLOG LEN",
    "(integer) 12",
    "127.0.0.1:6379>",
])

print("全部完成")
