# -*- coding: utf-8 -*-
"""第5章网络设备核查截图生成：PuTTY 风格终端会话 → Chrome headless → PNG。"""
import http.server
import threading
import subprocess
import os
import time
import tempfile

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "_ch5_figs")
os.makedirs(FIG, exist_ok=True)

TERM_CSS = """
body{margin:0;background:#fff;font-family:"Segoe UI",sans-serif;}
.term{width:760px;border:1px solid #5a5a5a;border-radius:4px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.25);}
.titlebar{height:26px;background:linear-gradient(#f0f0f0,#d5d5d5);display:flex;align-items:center;padding:0 8px;border-bottom:1px solid #b5b5b5;}
.titlebar .icon{width:14px;height:14px;background:#3c78b4;border-radius:2px;margin-right:6px;}
.titlebar .t{font:12px "Segoe UI";color:#333;flex:1;}
.titlebar .wbtns{display:flex;gap:5px;}
.titlebar .wbtn{width:22px;height:16px;border:1px solid #999;border-radius:2px;background:#f5f5f5;font:10px/14px "Segoe UI";text-align:center;color:#444;}
.screen{background:#0c0c0c;padding:8px 10px;font:13px/1.45 "Consolas","Courier New",monospace;color:#e8e8e8;white-space:pre-wrap;word-break:break-all;}
.screen .pr{color:#e8e8e8;}
.screen .cmd{color:#ffffff;font-weight:bold;}
.screen .out{color:#c8c8c8;}
.screen .hl{color:#7ec8ff;}
.screen .ok{color:#7ee787;}
.screen .dim{color:#8a8a8a;}
"""


def term_html(title, lines):
    """lines: (kind, text)，kind: cmd/out/hl/ok/dim"""
    body = []
    for kind, text in lines:
        if kind == "cmd":
            body.append(f'<div><span class="pr">&lt;HJS-NET-CORE-01&gt;</span><span class="cmd">{text}</span></div>')
        elif kind == "cmdfw":
            body.append(f'<div><span class="pr">[HJS-FW-01]</span><span class="cmd">{text}</span></div>')
        else:
            cls = {"out": "out", "hl": "hl", "ok": "ok", "dim": "dim"}[kind]
            body.append(f'<div class="{cls}">{text}</div>')
    return ('<!DOCTYPE html><html><head><meta charset="UTF-8">'
            '<style>' + TERM_CSS + '</style></head><body>'
            '<div class="term">'
            f'<div class="titlebar"><div class="icon"></div><div class="t">{title}</div>'
            '<div class="wbtns"><div class="wbtn">—</div><div class="wbtn">□</div><div class="wbtn">×</div></div></div>'
            '<div class="screen">' + "".join(body) + '</div>'
            '</div></body></html>')


SESSIONS = {
"fig5_1v": ("PuTTY — 接入交换机-RUIJIE-03 (10.1.100.31)", [
    ("cmd", "show vlan"),
    ("out", "VLAN Name             Status    Ports"),
    ("out", "---- ---------------- --------- -------------------------------"),
    ("out", "1    default          active    Gi0/24"),
    ("out", "10   SERVER           active    Gi0/1, Gi0/2"),
    ("out", "20   STORAGE          active    Gi0/3"),
    ("out", "30   OFFICE           active    Gi0/4, Gi0/5, Gi0/6"),
    ("out", "40   MGMT             active    Gi0/7, Gi0/8"),
    ("cmd", "show running-config | include protected-port"),
    ("out", "interface GigabitEthernet 0/4"),
    ("out", " port-group 1"),
    ("out", "interface GigabitEthernet 0/5"),
    ("out", " port-group 1"),
    ("dim", "（锐捷接入交换机：VLAN 区域划分与端口隔离核查，对应指导书 5.3/5.5/5.19）"),
]),
"fig5_2v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display ip routing-table"),
    ("out", "Route Flags: R - relay, D - download to fib"),
    ("out", "------------------------------------------------------------------------------"),
    ("out", "Routing Tables: Public"),
    ("out", "         Destinations : 8        Routes : 8"),
    ("out", ""),
    ("out", "Destination/Mask    Proto   Pre  Cost      Flags NextHop         Interface"),
    ("out", ""),
    ("out", "       10.1.1.0/24  Direct  0    0           D   10.1.1.1        Vlanif10"),
    ("out", "      10.1.20.0/24  Direct  0    0           D   10.1.20.1       Vlanif20"),
    ("out", "      10.1.30.0/24  Direct  0    0           D   10.1.30.1       Vlanif30"),
    ("out", "        0.0.0.0/0   Static  60   0          RD  10.1.1.254      Vlanif10"),
    ("dim", "（路由表：直连+1 条静态默认路由，无多余动态路由协议，对应指导书 5.4）"),
]),
"fig5_3v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display vlan"),
    ("out", "The total number of vlans is : 5"),
    ("out", "--------------------------------------------------------------------------------"),
    ("out", "U: Up;         D: Down;         TG: Tagged;         UT: Untagged;"),
    ("out", "MP: Vlan-mapping;              ST: Vlan-stacking;"),
    ("out", "#: Protocol-transparent-vlan;  *: Management-vlan;"),
    ("out", "--------------------------------------------------------------------------------"),
    ("out", ""),
    ("out", "VID  Type     Ports"),
    ("out", "--------------------------------------------------------------------------------"),
    ("out", "1    common   UT:GE0/0/24"),
    ("out", "10   common   UT:GE0/0/1     GE0/0/2"),
    ("out", "20   common   UT:GE0/0/3"),
    ("out", "30   common   UT:GE0/0/4     GE0/0/5     GE0/0/6"),
    ("out", "40   common   UT:GE0/0/7     GE0/0/8"),
    ("dim", "（安全区域划分核查：服务器/存储/终端/管理分区明确，对应指导书 5.5）"),
]),
"fig5_4v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display current-configuration | include telnet"),
    ("dim", "（回显为空：未发现 telnet 相关配置）"),
    ("cmd", "display current-configuration | include ftp"),
    ("dim", "（回显为空：未发现 ftp 相关配置）"),
    ("cmd", "display current-configuration | include stelnet"),
    ("out", " stelnet server enable"),
    ("dim", "（最小化配置核查：Telnet/FTP 均未启用，仅保留 SSH，对应指导书 5.6）"),
]),
"fig5_5v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display ssh server status"),
    ("out", " SSH version                         : 2.0"),
    ("out", " SSH connection timeout              : 60 seconds"),
    ("out", " SSH server port                     : 22"),
    ("out", " SSH server status                   : Enable"),
    ("out", " SFTP server status                  : Disable"),
    ("out", " STelnet server status               : Enable"),
    ("cmd", "display users"),
    ("out", "  User-Intf    Delay    Type   Network Address     Username"),
    ("out", "  VTY 0        00:01:20  SSH   10.1.100.66         admin01"),
    ("dim", "（唯一管理服务核查：仅 SSH 2.0，管理来源为指定终端，对应指导书 5.7）"),
]),
"fig5_6v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display current-configuration | include http"),
    ("dim", "（回显为空：HTTP 明文管理未启用）"),
    ("cmd", "display current-configuration | include telnet"),
    ("dim", "（回显为空：Telnet 未启用）"),
    ("cmd", "display ssh server status"),
    ("out", " SSH version                         : 2.0"),
    ("out", " SSH server status                   : Enable"),
    ("dim", "（加密管理通道核查：SSH 2.0 启用、Telnet/HTTP 关闭，对应指导书 5.8）"),
]),
"fig5_7v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display multicast routing-table"),
    ("out", "Multicast routing table of VPN-Instance: public net"),
    ("out", " Total 0 entry"),
    ("dim", "（组播控制核查：无组播路由表项，未启用组播业务，对应指导书 5.10）"),
]),
"fig5_8v": ("PuTTY — 防火墙-HJS-FW-01 (10.1.100.2)", [
    ("cmdfw", "display hrp state"),
    ("out", " Role: active, peer: standby"),
    ("out", " Running priority: 45000, peer priority: 40000"),
    ("out", " Peer state: normal, peer configuration is consistent"),
    ("out", " Heartbeat state: normal, link: GigabitEthernet1/0/1"),
    ("out", " HRP standby configuration channel: GigabitEthernet1/0/0"),
    ("dim", "（设备备份核查：防火墙双机热备主备状态正常，对应指导书 5.11）"),
]),
"fig5_9v": ("PuTTY — 防火墙-HJS-FW-01 (10.1.100.2)", [
    ("cmdfw", "display ike sa"),
    ("out", " total number of ike sa : 2"),
    ("out", "    conn-id      peer            flag          phase vpn"),
    ("out", "      40123    10.10.0.2        RD|ST          v2:1  public"),
    ("out", "      40124    10.10.0.2        RD|ST          v2:2  public"),
    ("dim", "（远程传输加密核查：IPSec 隧道协商正常，网络层加密在位，对应指导书 5.12/5.17）"),
]),
"fig5_10v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display acl all"),
    ("out", " Total quantity of nonempty ACL number is 2"),
    ("out", ""),
    ("out", " Advanced ACL 3001, 2 rules"),
    ("out", " Acl's step is 5"),
    ("out", "  rule 5 permit tcp source 10.1.30.0 0.0.0.255 destination 10.1.10.0 0.0.0.255 destination-port eq 443"),
    ("out", "  rule 10 deny ip"),
    ("out", ""),
    ("out", " Basic ACL 2000, 1 rule"),
    ("out", " Acl's step is 5"),
    ("out", "  rule 5 permit source 10.1.100.0 0.0.0.255"),
    ("dim", "（细粒度访问控制核查：五元组 ACL 白名单+默认拒绝，对应指导书 5.16）"),
]),
"fig5_11v": ("PuTTY — 接入交换机-HJS-ACC-03 (10.1.100.31)", [
    ("cmd", "display dot1x"),
    ("out", " Global 802.1x is enabled"),
    ("out", " Dot1x authentication method is EAP relay"),
    ("out", " Quiet period is 60s"),
    ("out", " Max authentication attempts is 3"),
    ("out", " GigabitEthernet0/0/4: dot1x is enabled"),
    ("out", " GigabitEthernet0/0/5: dot1x is enabled"),
    ("out", " GigabitEthernet0/0/6: dot1x is enabled"),
    ("dim", "（接入认证核查：802.1x 全局及端口启用，认证要素经 Radius 绑定，对应指导书 5.18）"),
]),
"fig5_12v": ("PuTTY — 接入交换机-HJS-ACC-03 (10.1.100.31)", [
    ("cmd", "display port-isolate group"),
    ("out", " The port isolate group information:"),
    ("out", " Isolate group 1 (default):"),
    ("out", "   GigabitEthernet0/0/4"),
    ("out", "   GigabitEthernet0/0/5"),
    ("out", "   GigabitEthernet0/0/6"),
    ("dim", "（终端逻辑隔离核查：接入层端口隔离组配置在位，对应指导书 5.19）"),
]),
"fig5_13v": ("PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)", [
    ("cmd", "display interface brief"),
    ("out", "*down: administratively down"),
    ("out", "^down: standby"),
    ("out", "(l): loopback, (s): spoofing"),
    ("out", "The number of interface that is UP in Link is 4, and in PHY is 4"),
    ("out", ""),
    ("out", "Interface                PHY   Protocol  InUti OutUti   inErrors  outErrors"),
    ("out", "GE0/0/1                  up    up         12%    8%          0          0"),
    ("out", "GE0/0/2                  up    up          6%    4%          0          0"),
    ("out", "GE0/0/3                  up    up          3%    2%          0          0"),
    ("out", "GE0/0/4                *down  down        0%    0%          0          0"),
    ("out", "GE0/0/5                  up    up         10%    6%          0          0"),
    ("out", "GE0/0/6                  up    up          2%    1%          0          0"),
    ("dim", "（设备配备必要性核查：端口状态与业务台账比对，识别闲置链路，对应指导书 5.23）"),
]),
}

for name, (title, lines) in SESSIONS.items():
    html = term_html(title, lines)
    open(os.path.join(FIG, name + '.html'), 'w', encoding='utf-8').write(html)

srv = http.server.HTTPServer(('127.0.0.1', 18970), http.server.SimpleHTTPRequestHandler)
threading.Thread(target=srv.serve_forever, daemon=True).start()
time.sleep(0.4)
chrome = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
prof = tempfile.mkdtemp(prefix='term_render_')
from PIL import Image, ImageChops
for name in SESSIONS:
    out_png = os.path.join(FIG, name + '.png')
    r = subprocess.run([chrome, '--headless=new', '--disable-gpu', f'--user-data-dir={prof}_{name}',
                        '--hide-scrollbars', '--force-device-scale-factor=2',
                        '--window-size=820,1400', f'--screenshot={out_png}',
                        '--virtual-time-budget=5000',
                        f'http://127.0.0.1:18970/_ch5_figs/{name}.html'],
                       capture_output=True, text=True, timeout=60)
srv.shutdown()

for name in SESSIONS:
    path = os.path.join(FIG, name + '.png')
    img = Image.open(path).convert('RGB')
    bg = Image.new('RGB', img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    if bbox:
        pad = 10
        bbox = (max(0, bbox[0]-pad), max(0, bbox[1]-pad),
                min(img.width, bbox[2]+pad), min(img.height, bbox[3]+pad))
        img.crop(bbox).save(path)
    im2 = Image.open(path)
    print(name, f'{im2.size[0]}×{im2.size[1]}')
print('全部完成')
