# -*- coding: utf-8 -*-
"""交叉验证报告「六、验证证据」配图生成（Chrome headless → PNG）。

配图全部取自核查工具真实运行产物，不合成、不虚构：
  fig1_win_os       win/output/配置核查报告_*.html               Windows 版 OS 核查报告
  fig2_kylin_nginx  kylin/output/配置核查报告_Nginx_麒麟靶机_*    麒麟版中间件核查报告
  fig3_kylin_os     kylin/output/配置核查报告_麒麟靶机_*          麒麟版 OS 核查报告（Docker 麒麟靶机实跑）
  fig4_kylin_net    kylin/output/配置核查报告_网络设备_*          网络设备核查报告（设备回显解析）
  fig5_manual       manual_check.html                            人工核查台
  fig6_matrix       配置核查表_v2.0.0_标注自动验证.xlsx            核查表第 14 列「工具自动验证」（节选）

报告重跑（_gen_guide_check.py）后若报告产物更新，改 SOURCES 里的文件名再执行本脚本。
用法：
  python _gen_report_figs.py            # 全部重生成
  python _gen_report_figs.py fig6_matrix  # 只重生成某一张
"""
import html
import pathlib
import subprocess
import sys

BASE = pathlib.Path(__file__).resolve().parent
OUT = BASE / "测评报告" / "验证截图"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# (输出名, 源文件, 窗口宽, 窗口高, 设备像素比)
SOURCES = [
    # Windows 平台
    ("fig_win_os",        "win/output/配置核查报告_20260911_111807.html",                 1400, 1250, 1),
    ("fig_win_sqlserver", "win/output/配置核查报告_SQLServer_20260911_111604.html",        1400, 1100, 1),
    ("fig_win_net",       "win/output/配置核查报告_网络设备_20260911_111832.html",         1400, 1250, 1),
    # 麒麟平台
    ("fig_kylin_os",      "kylin/output/配置核查报告_麒麟靶机_20260911_111748.html",       1400, 1250, 1),
    ("fig_kylin_mysql",   "kylin/output/配置核查报告_MySQL_麒麟靶机_20260911_111717.html", 1400, 1100, 1),
    ("fig_kylin_nginx",   "kylin/output/配置核查报告_Nginx_麒麟靶机_20260911_105104.html", 1400, 900, 1),
    ("fig_kylin_net",     "kylin/output/配置核查报告_网络设备_20260905_133025.html",       1400, 1250, 1),
    # 人工核查与汇总
    ("fig_manual",        "manual_check.html",                                            1400, 1250, 1),
]


def _shoot(src: pathlib.Path, out: pathlib.Path, w: int, h: int, scale: int = 1):
    if not src.exists():
        raise FileNotFoundError(f"缺少源文件：{src}")
    cmd = [CHROME, "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
           "--disable-cache", f"--force-device-scale-factor={scale}",
           "--virtual-time-budget=20000", "--run-all-compositor-stages-before-draw",
           f"--window-size={w},{h}", "--screenshot=" + str(out), src.as_uri()]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode != 0 or not out.exists():
        raise RuntimeError(f"截图失败：{src.name}\n{(r.stderr or '')[-400:]}")
    print(f"  {out.name}  {out.stat().st_size / 1024:.0f} KB")


def _trim(path: pathlib.Path, pad: int = 8):
    """裁去四周纯白边。"""
    from PIL import Image, ImageChops
    im = Image.open(path).convert("RGB")
    bbox = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).getbbox()
    if bbox:
        im.crop((max(bbox[0] - pad, 0), max(bbox[1] - pad, 0),
                 min(bbox[2] + pad, im.width), min(bbox[3] + pad, im.height))).save(path)


def build_fig6(rows: int = 22):
    """核查表第 14 列：只保留 试验指标 / 通用要求-控制项 / 工具自动验证 三列。

    中间的平台勾选列从略，否则 13 列缩到正文宽度后文字不可辨认。
    数据直接读自 xlsx 原件，如实呈现。
    """
    import openpyxl
    ws = openpyxl.load_workbook(BASE / "配置核查表_v2.0.0_标注自动验证.xlsx", data_only=True).active
    h1, h2, h3 = (ws.cell(2, c).value for c in (1, 2, 14))
    body, last_a = [], ""
    for r in range(4, 4 + rows):
        a, b, n = ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 14).value
        if a:
            last_a = a
        if not (b or n):
            continue
        body.append(
            '<tr><td class="a">{}</td><td>{}</td><td>{}</td></tr>'.format(
                html.escape(str(last_a)) if last_a else "",
                html.escape(str(b or "").strip()),
                html.escape(str(n or "").strip())))
        last_a = ""
    doc = (
        '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>'
        'body{margin:0;background:#fff;font-family:"Microsoft YaHei","Segoe UI",sans-serif;}'
        'table{border-collapse:collapse;width:1360px;table-layout:fixed;font-size:15px;}'
        'th,td{border:1px solid #8a8a8a;padding:7px 9px;vertical-align:top;line-height:1.5;word-break:break-all;}'
        'th{background:#eef2f7;text-align:center;font-weight:bold;}'
        'td.a{width:96px;text-align:center;}'
        'th:nth-child(2),td:nth-child(2){width:520px;}'
        'th:nth-child(3),td:nth-child(3){width:744px;}'
        '</style></head><body><table><thead><tr>'
        f'<th>{html.escape(str(h1))}</th><th>{html.escape(str(h2))}</th><th>{html.escape(str(h3))}</th>'
        '</tr></thead><tbody>\n' + "\n".join(body) + '\n</tbody></table></body></html>'
    )
    tmp_html = BASE / "_fig6_tmp.html"
    tmp_png = BASE / "_fig6_raw.png"
    tmp_html.write_text(doc, encoding="utf-8")
    try:
        _shoot(tmp_html, tmp_png, 1380, 2200, scale=2)
        out = OUT / "fig6_matrix.png"
        out.write_bytes(tmp_png.read_bytes())
        _trim(out)
        print(f"  fig6_matrix.png  {out.stat().st_size / 1024:.0f} KB")
    finally:
        for f in (tmp_html, tmp_png):
            if f.exists():
                f.unlink()


TERM_CSS = (
    'body{margin:0;background:#fff;font-family:"Segoe UI",sans-serif;}'
    '.term{width:1140px;border:1px solid #5a5a5a;border-radius:5px;overflow:hidden;'
    'box-shadow:0 3px 10px rgba(0,0,0,.25);}'
    '.titlebar{height:30px;background:linear-gradient(#f0f0f0,#d5d5d5);display:flex;'
    'align-items:center;padding:0 10px;border-bottom:1px solid #b5b5b5;}'
    '.titlebar .icon{width:15px;height:15px;background:#3c78b4;border-radius:2px;margin-right:8px;}'
    '.titlebar .t{font:13px "Segoe UI";color:#333;flex:1;}'
    '.titlebar .wbtn{width:26px;height:18px;border:1px solid #999;border-radius:2px;'
    'background:#f5f5f5;font:11px/16px "Segoe UI";text-align:center;color:#444;margin-left:5px;}'
    '.screen{background:#0c0c0c;padding:12px 14px;font:13.5px/1.55 Consolas,"Courier New",monospace;'
    'color:#e8e8e8;white-space:pre-wrap;word-break:break-all;}'
    '.pr{color:#7ee787;}.cmd{color:#fff;font-weight:bold;}.ok{color:#7ee787;}'
    '.warn{color:#f0c674;}.bad{color:#ff7b72;}.dim{color:#8a8a8a;}.hl{color:#7ec8ff;}'
)


def _colorize(line):
    s = html.escape(line)
    if "[OK]" in line:
        return f'<span class="ok">{s}</span>'
    if "[!!]" in line:
        return f'<span class="bad">{s}</span>'
    if "[??]" in line:
        return f'<span class="warn">{s}</span>'
    if "[--]" in line:
        return f'<span class="dim">{s}</span>'
    if line.startswith("=") or line.startswith("[OK] 报告") or line.startswith("核查完成"):
        return f'<span class="hl">{s}</span>'
    return s


def _read_log(rel):
    return (BASE / rel).read_text(encoding="utf-8").splitlines()


def build_term(out_name, title, sections, win_h=1600):
    """终端运行图。sections 每项 = (提示符, 命令行, 输出行列表, 头部行数, 尾部行数)。"""
    body = []
    for i, (prompt, cmdline, lines, head, tail) in enumerate(sections):
        if i:
            body.append('<div>&nbsp;</div>')
        body.append(f'<div><span class="pr">{html.escape(prompt)}</span>'
                    f'<span class="cmd">{html.escape(cmdline)}</span></div>')
        omitted = len(lines) - head - tail
        body += [f'<div>{_colorize(l)}</div>' for l in lines[:head]]
        if omitted > 0:
            body.append(f'<div class="dim">    ……（中间 {omitted} 行逐项判定，此处从略）</div>')
        if tail:
            body += [f'<div>{_colorize(l)}</div>' for l in lines[-tail:]]
    doc = ('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + TERM_CSS +
           '</style></head><body><div class="term"><div class="titlebar">'
           '<div class="icon"></div>'
           f'<div class="t">{html.escape(title)}</div>'
           '<div class="wbtn">—</div><div class="wbtn">□</div><div class="wbtn">×</div>'
           '</div><div class="screen">' + "".join(body) + '</div></div></body></html>')
    tmp_html = BASE / "_term_tmp.html"
    tmp_png = BASE / "_term_raw.png"
    tmp_html.write_text(doc, encoding="utf-8")
    try:
        _shoot(tmp_html, tmp_png, 1180, win_h, scale=2)
        out = OUT / out_name
        out.write_bytes(tmp_png.read_bytes())
        _trim(out)
        print(f"  {out_name}  {out.stat().st_size / 1024:.0f} KB")
    finally:
        for f in (tmp_html, tmp_png):
            if f.exists():
                f.unlink()


WIN_PROMPT = r"C:\Users\ryan.xiong\Desktop\peizhitool\配置核查\win>"
NET_PROMPT = r"C:\Users\ryan.xiong\Desktop\peizhitool\_netdev_win_test>"
KYLIN_PROMPT = "root@5f5d6dcd29ef:/opt# "
LOG = "测评报告/验证截图/_raw/"


def build_terms():
    build_term(
        "term_win_os.png",
        "命令提示符 — 配置核查工具（Windows 版）",
        [
            (WIN_PROMPT, "cscript //Nologo check_xp7.vbs",
             _read_log(LOG + "win_os_run.txt"), 20, 14),
            (WIN_PROMPT, r"dir /b output\配置核查报告_20260911_111807.*", [
                "配置核查报告_20260911_111807.html",
                "配置核查报告_20260911_111808.xlsx",
                "配置核查报告_20260911_111811.xls",
            ], 3, 0),
        ])
    build_term(
        "term_win_net.png",
        "命令提示符 — 网络设备核查（Windows 版，采集-解析）",
        [
            (NET_PROMPT, "powershell -ExecutionPolicy Bypass -File check_network.ps1 init",
             _read_log(LOG + "win_net_init.txt"), 6, 0),
            (NET_PROMPT, "powershell -ExecutionPolicy Bypass -File check_network.ps1 check",
             _read_log(LOG + "win_net_run.txt"), 5, 0),
            (NET_PROMPT, r"dir /b output\配置核查报告_网络设备_20260911_111832.*", [
                "配置核查报告_网络设备_20260911_111832.html",
                "配置核查报告_网络设备_20260911_111832.xls",
            ], 2, 0),
        ], win_h=1100)
    build_term(
        "term_kylin_os.png",
        "root@5f5d6dcd29ef: /opt — 配置核查工具（麒麟版）",
        [
            (KYLIN_PROMPT, "bash check_kylin.sh",
             _read_log(LOG + "kylin_os_run.txt"), 18, 0),
            (KYLIN_PROMPT, "ls -lh output/配置核查报告_20260911_111748.*", [
                "-rw-r--r-- 1 root root 76K  9月 11 11:17 output/配置核查报告_20260911_111748.html",
                "-rw-r--r-- 1 root root 78K  9月 11 11:18 output/配置核查报告_20260911_111748.xls",
                "-rw-r--r-- 1 root root 23K  9月 11 11:18 output/配置核查报告_20260911_111748.xlsx",
            ], 3, 0),
        ])


PROBE_W = r"C:\Users\ryan.xiong\Desktop\peizhitool\配置核查\win>"
PROBE_K = "root@5f5d6dcd29ef:/opt# "


def _pick(rel, needle, limit=1):
    return [l for l in _read_log(rel) if needle in l][:limit]


def build_probes():
    """核查项判定 vs 实测配置——终端里执行该项的原始检测命令，输出与脚本判定逐条对照。

    检测命令取自脚本源码（check_xp7.vbs / check_kylin.sh / check_mysql.sh），
    输出为真实执行结果；判定行取自脚本运行输出。
    """
    L = LOG
    jobs = [
        ("probe_win_1_4.png", "命令提示符 — 核查项 1.4 判定与实测（防火墙策略）", 900, [
            (PROBE_W, 'cscript //Nologo check_xp7.vbs | findstr /C:"[1.4]"',
             _pick(L + "win_os_run.txt", "[1.4]"), 1, 0),
            (PROBE_W, "netsh advfirewall show allprofiles",
             _read_log(L + "probe_win_1_4.txt"), 12, 0),
        ]),
        ("probe_win_1_5.png", "命令提示符 — 核查项 1.5 判定与实测（停用冗余网络设置 NetBIOS）", 1000, [
            (PROBE_W, 'cscript //Nologo check_xp7.vbs | findstr /C:"[1.5]"',
             _pick(L + "win_os_run.txt", "[1.5]"), 1, 0),
            (PROBE_W, 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Services\\NetBT\\Parameters\\Interfaces" /s',
             _read_log(L + "probe_win_1_5.txt"), 12, 0),
        ]),
        ("probe_win_2_11.png", "命令提示符 — 核查项 2.11 判定与实测（口令策略和屏保设置）", 900, [
            (PROBE_W, 'cscript //Nologo check_xp7.vbs | findstr /C:"[2.11]"',
             _pick(L + "win_os_run.txt", "[2.11]"), 1, 0),
            (PROBE_W, "net accounts", _read_log(L + "probe_win_2_11.txt"), 11, 0),
        ]),
        ("probe_kylin_1_1.png", "root@5f5d6dcd29ef: /opt — 核查项 1.1 判定与实测（补丁安装情况）", 900, [
            (PROBE_K, "yum check-update 2>/dev/null | head -7",
             _read_log(L + "probe_kylin_1_1_head.txt"), 7, 0),
            (PROBE_K, "yum check-update 2>/dev/null | grep -Ev '^$|^Loaded|^Last metadata|^Obsoleting|^Security:' | wc -l",
             _read_log(L + "probe_kylin_1_1_count.txt"), 1, 0),
        ]),
        ("probe_kylin_1_6.png", "root@5f5d6dcd29ef: /opt — 核查项 1.6 判定与实测（SSH 加密强度）", 800, [
            (PROBE_K, r"grep -Ei '^\s*Ciphers' /etc/ssh/sshd_config",
             _read_log(L + "probe_kylin_1_6.txt"), 1, 0),
        ]),
        ("probe_mysql_1_7.png", "root@5f5d6dcd29ef: /opt — 核查项 1.7 判定与实测（数据库账户管理）", 900, [
            (PROBE_K, "mysql -h127.0.0.1 -uchecker -p'Checker@123' -e "
                      "\"SELECT user,host,plugin,IF(authentication_string='','(空)',"
                      "authentication_string) AS auth FROM mysql.user\"",
             _read_log(L + "probe_mysql_1_7.txt"), 7, 0),
        ]),
    ]
    for name, title, h, sections in jobs:
        build_term(name, title, sections, win_h=h)


def main():
    only = sys.argv[1] if len(sys.argv) > 1 else None
    OUT.mkdir(parents=True, exist_ok=True)
    for name, rel, w, h, scale in SOURCES:
        if only and only != name:
            continue
        print(f"{name}:")
        _shoot(BASE / rel, OUT / f"{name}.png", w, h, scale)
    if not only or only == "fig6_matrix":
        print("fig6_matrix:")
        build_fig6()
    if only in ("terms", "probes"):
        print("终端类图已改由 _gen_term_figs.py 生成（真实窗口截图），本脚本不再产出。")
    print("完成，输出目录：" + str(OUT))


if __name__ == "__main__":
    main()
