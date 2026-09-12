# -*- coding: utf-8 -*-
r"""第 2 章缺图小节配图生成并插入。

五张图（_tmp_figs/）：
  gen_ch2_shadow.png    2.4.3   cat /etc/shadow 比对用户口令（麒麟终端）
  gen_ch2_gpedit.png    2.8.1   本地组策略编辑器：账户策略（窗口示意）
  gen_ch2_regquery.png  2.11.1  reg query 查屏保策略（cmd）
  gen_ch2_auditpol.png  2.15.1  auditpol + net accounts（cmd）
  gen_ch2_kylin.png     2.15.2  firewalld / auditctl / login.defs（麒麟终端）

用法：
  python _gen_ch2_figs.py            # 只生成图片
  python _gen_ch2_figs.py --apply    # 插入 docx（先自行备份）
"""
import os
import subprocess
import sys

from PIL import Image, ImageChops
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm

HERE = os.path.dirname(os.path.abspath(__file__))
FIGS = os.path.join(HERE, "_tmp_figs")
PATH = os.path.join(HERE, "配置核查作业指导书_v2.2.docx")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
APPLY = "--apply" in sys.argv
DPI = 110.0

TERM_CSS = ('body{margin:0;background:#fff;}'
            'pre{font:15px/1.9 "Courier New","Consolas",monospace;color:#000;'
            'margin:0;padding:6px 10px;white-space:pre;}')


def shoot(html_path, png_path, w, h):
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
                    "--hide-scrollbars", "--force-device-scale-factor=1",
                    "--virtual-time-budget=3000", "--window-size=%d,%d" % (w, h),
                    "--screenshot=" + png_path, "file:///" + html_path.replace("\\", "/")],
                   capture_output=True, text=True)


def save_trim(png_path, pad=8):
    im = Image.open(png_path).convert("RGB")
    bg = Image.new("RGB", im.size, (255, 255, 255))
    bbox = ImageChops.difference(im, bg).getbbox()
    if bbox:
        im.crop((max(bbox[0] - pad, 0), max(bbox[1] - pad, 0),
                 min(bbox[2] + pad, im.width), min(bbox[3] + pad, im.height))).save(png_path)
    return Image.open(png_path).size


def term(name, lines):
    body = "\n".join(lines).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    html_path = os.path.join(FIGS, "_t.html")
    png_path = os.path.join(FIGS, name + ".png")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{margin:0;background:#fff;}'
                'pre{%s}</style></head><body><pre>%s</pre></body></html>' % (TERM_CSS, body))
    shoot(html_path, png_path, 1100, 800)
    if os.path.exists(html_path):
        os.remove(html_path)
    print(name, save_trim(png_path))


GPEDIT_CSS = """
body{margin:0;background:#fff;font-family:"Microsoft YaHei",sans-serif;}
.win{width:660px;border:1px solid #7a7a7a;border-radius:3px;overflow:hidden;box-shadow:2px 3px 8px rgba(0,0,0,.25);}
.tbar{background:linear-gradient(#f5f6f7,#dfe1e3);padding:5px 10px;font:13px "Microsoft YaHei";color:#333;border-bottom:1px solid #bbb;}
.menu{padding:3px 10px;font:12px "Microsoft YaHei";color:#444;border-bottom:1px solid #ddd;}
.flex{display:flex;}
.tree{width:250px;font:12.5px/1.9 "Microsoft YaHei";color:#222;padding:8px 6px;white-space:pre;border-right:1px solid #d9dbde;}
.main{flex:1;padding:8px 10px;}
.main table{width:100%;border-collapse:collapse;font:12px "Microsoft YaHei";}
.main th{background:#eef1f4;border:1px solid #ccd0d4;padding:4px 8px;text-align:left;}
.main td{border:1px solid #ccd0d4;padding:4px 8px;}
.sbar{border-top:1px solid #d9dbde;padding:3px 10px;font:12px "Microsoft YaHei";color:#666;}
"""


def gpedit(name):
    html = ('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>%s</style></head><body>'
            '<div class="win"><div class="tbar">本地组策略编辑器</div>'
            '<div class="menu">文件(F)　操作(A)　查看(V)　帮助(H)</div>'
            '<div class="flex"><div class="tree">'
            '<b>本地计算机 策略</b>\n'
            '▾ 计算机配置\n'
            '　▾ Windows 设置\n'
            '　　▾ 安全设置\n'
            '　　　▸ <b>账户策略</b>\n'
            '　　　　▸ 密码策略\n'
            '　　　　▸ 账户锁定策略\n'
            '　　　▸ 本地策略\n'
            '　　　▸ 网络列表管理策略\n'
            '</div><div class="main"><table>'
            '<tr><th>策略</th><th>状态</th></tr>'
            '<tr><td>密码必须符合复杂性要求</td><td>已启用</td></tr>'
            '<tr><td>密码长度最小值</td><td>10 个字符</td></tr>'
            '<tr><td>密码最长使用期限</td><td>30 天</td></tr>'
            '<tr><td>强制密码历史</td><td>5 个记住的密码</td></tr>'
            '<tr><td>账户锁定阈值</td><td>5 次无效登录</td></tr>'
            '</table></div></div>'
            '<div class="sbar">本地组策略编辑器</div></div></body></html>' % GPEDIT_CSS)
    html_path = os.path.join(FIGS, "_t.html")
    png_path = os.path.join(FIGS, name + ".png")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    shoot(html_path, png_path, 760, 700)
    if os.path.exists(html_path):
        os.remove(html_path)
    print(name, save_trim(png_path, pad=0))


def main():
    os.makedirs(FIGS, exist_ok=True)

    # 2.4.3 口令互不相同
    term("gen_ch2_shadow", [
        "[root@kylin-server ~]# cat /etc/shadow",
        "root:$6$Xk9s…hashed…:$6$Qw2r…:19923:0:99999:7:::",
        "daemon:*:19023:0:99999:7:::",
        "kylin:$6$Pz5t…hashed…:$6$Mn8v…:19961:0:30:7:::",
        "audit:$6$Lm4n…hashed…:$6$Pq7r…:19961:0:30:7:::",
        "",
        "# 比对第二字段：各账户哈希均不同 → 口令互不相同",
        "[root@kylin-server ~]# passwd -S kylin",
        "kylin PS 2026-07-15 0 30 7 -1 (已设置口令，可登录)",
    ])

    # 2.11.1 reg query 屏保策略
    term("gen_ch2_regquery", [
        r'C:\Users\zhang_wei> reg query "HKCU\Control Panel\Desktop" /v ScreenSaverIsSecure',
        r'HKEY_CURRENT_USER\Control Panel\Desktop',
        "    ScreenSaverIsSecure    REG_SZ    1",
        "",
        r'C:\Users\zhang_wei> reg query "HKCU\Control Panel\Desktop" /v ScreenSaveTimeOut',
        r'HKEY_CURRENT_USER\Control Panel\Desktop',
        "    ScreenSaveTimeOut    REG_SZ    300",
        "",
        r'C:\Users\zhang_wei> reg query "HKCU\Control Panel\Desktop" /v ScreenSaveActive',
        r'HKEY_CURRENT_USER\Control Panel\Desktop',
        "    ScreenSaveActive    REG_SZ    1",
    ])

    # 2.15.1 auditpol + net accounts
    term("gen_ch2_auditpol", [
        r'C:\Windows\system32> auditpol /get /category:*',
        "系统审核策略",
        "类别/子类别                       设置",
        "登录/注销                          成功和失败",
        "  登录                             成功和失败",
        "账户管理                           成功和失败",
        "策略更改                           成功和失败",
        "特权使用                           无审核",
        "",
        r'C:\Windows\system32> net accounts',
        "密码最短使用期限(天):                    1",
        "密码最长使用期限(天):                    30",
        "密码长度最小值:                          10",
        "锁定阈值:                                5",
    ])

    # 2.15.2 麒麟安全策略
    term("gen_ch2_kylin", [
        "[root@kylin-server ~]# systemctl status firewalld | head -3",
        "● firewalld.service - firewalld - dynamic firewall daemon",
        "   Active: active (running) since 五 2026-07-10 09:30:12 CST; Enabled",
        "[root@kylin-server ~]# auditctl -l",
        "-w /etc/passwd -p wa -k identity",
        "-a always,exit -F arch=b64 -S execve -k auditcmd",
        "[root@kylin-server ~]# grep PASS_MAX_DAYS /etc/login.defs",
        "PASS_MAX_DAYS   30",
    ])

    # 2.8.1 gpedit
    gpedit("gen_ch2_gpedit")

    if not APPLY:
        print("（图片已生成；加 --apply 插入文档）")
        return 0

    # ---------------- 插入 ----------------
    d = Document(PATH)
    paras = d.paragraphs

    JOBS = [
        ("输入“cat /etc/shadow”命令", "gen_ch2_shadow.png"),        # 2.4.3
        ("网络位置配置为“受管理”", "gen_ch2_gpedit.png"),            # 2.8.1 末段后
        ("reg query \"HKCU\\Control Panel\\Desktop\" /v ScreenSaveActive", "gen_ch2_regquery.png"),  # 2.11.1
        ("核查本地安全策略。使用win+R组合键打开运行框，输入secpol.msc", "gen_ch2_auditpol.png"),  # 2.15.1 末段后
        ("重点查看PASS_MAX_DAYS", "gen_ch2_kylin.png"),              # 2.15.2 末段后
    ]
    for prefix, fig in JOBS:
        anchor = None
        for p in paras:
            if p.text.strip().startswith(prefix) or prefix in p.text:
                anchor = p
                break
        if anchor is None:
            print("!! 未找到锚点:", prefix[:30])
            continue
        ps = d.paragraphs
        nxt = None
        for i, p in enumerate(ps):
            if p._p is anchor._p:
                nxt = ps[i + 1] if i + 1 < len(ps) else None
                break
        if nxt is not None and nxt._p.findall('.//' + qn('a:blip')) and not nxt.text.strip():
            print("已有图，跳过:", fig)
            continue
        src = os.path.join(FIGS, fig)
        w, h = Image.open(src).size
        w_cm = min(14.64, w / DPI * 2.54)
        img = d.add_paragraph()
        img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        img.paragraph_format.first_line_indent = None
        img.add_run().add_picture(src, width=Cm(w_cm))
        anchor._p.addnext(img._p)
        print("已插入 %s（%.2fcm）于 %r 后" % (fig, w_cm, anchor.text.strip()[:24]))

    d.save(PATH)
    print("已保存", PATH)
    return 0


if __name__ == "__main__":
    sys.exit(main())
