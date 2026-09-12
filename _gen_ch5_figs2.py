# -*- coding: utf-8 -*-
"""第 5 章各小节配图批量生成（PuTTY 风格终端图，风格与 _gen_ch5_terms.py 一致）。

每节 2-3 张，插入到各小节末尾。新图命名 fig5b_*，与既有 fig5_* 不冲突。
用法：
  python _gen_ch5_figs2.py            # 只生成图片到 _ch5_figs2/
  python _gen_ch5_figs2.py --apply    # 插入 docx（先自行备份）
"""
import os
import subprocess
import re
import sys

from docx import Document
from docx.shared import Cm
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(HERE, "_ch5_figs2")
PATH = os.path.join(HERE, "配置核查作业指导书_v2.2.docx")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
APPLY = "--apply" in sys.argv

TERM_CSS = """
body{margin:0;background:#fff;font-family:"Segoe UI",sans-serif;}
.term{width:760px;border:1px solid #5a5a5a;border-radius:4px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.25);}
.titlebar{height:26px;background:linear-gradient(#f0f0f0,#d5d5d5);display:flex;align-items:center;padding:0 8px;border-bottom:1px solid #b5b5b5;}
.titlebar .icon{width:14px;height:14px;background:#3c78b4;border-radius:2px;margin-right:6px;}
.titlebar .t{font:12px "Segoe UI";color:#333;flex:1;}
.titlebar .wbtns{display:flex;gap:5px;}
.titlebar .wbtn{width:22px;height:16px;border:1px solid #999;border-radius:2px;background:#f5f5f5;font:10px/14px "Segoe UI";text-align:center;color:#444;}
.screen{background:#0c0c0c;padding:8px 10px;font:13px/1.45 "Consolas","Courier New",monospace;color:#e8e8e8;white-space:pre-wrap;word-break:break-all;}
.screen .cmd{color:#ffffff;font-weight:bold;}
.screen .out{color:#c8c8c8;}
.screen .hl{color:#7ec8ff;}
.screen .ok{color:#7ee787;}
.screen .dim{color:#8a8a8a;}
"""

PROMPTS = {
    "hw": "<HJS-CORE-01>",
    "fw": "[HJS-FW-01]",
    "rj": "<RUIJIE-ACC-03>",
    "h3c": "<HJS-H3C-01>",
    "ac": "<HJS-AC-01>",
}


def term_html(title, lines):
    """lines: (kind, text)；kind: cmd/out/hl/ok/dim（提示符按标题里的设备推断）"""
    if "RUIJIE" in title:
        pr = PROMPTS["rj"]
    elif "无线控制器" in title:
        pr = PROMPTS["ac"]
    elif "防火墙" in title:
        pr = PROMPTS["fw"]
    elif "H3C" in title:
        pr = PROMPTS["h3c"]
    else:
        pr = PROMPTS["hw"]
    body = []
    for kind, text in lines:
        if kind == "cmd":
            body.append('<div><span class="out">%s</span><span class="cmd">%s</span></div>' % (pr, text))
        else:
            body.append('<div class="%s">%s</div>' % (kind, text))
    return ('<!DOCTYPE html><html><head><meta charset="UTF-8">'
            '<style>' + TERM_CSS + '</style></head><body>'
            '<div class="term">'
            '<div class="titlebar"><div class="icon"></div><div class="t">%s</div>'
            '<div class="wbtns"><div class="wbtn">—</div><div class="wbtn">□</div><div class="wbtn">×</div></div></div>'
            '<div class="screen">%s</div></div></body></html>' % (title, "".join(body)))


# ---------------- 每节配图定义：sec -> [(名称, 标题, 提示符, 行)] ----------------
HW = "PuTTY — 核心交换机-HJS-CORE-01 (10.1.100.10)"
FW = "PuTTY — 防火墙-HJS-FW-01 (10.1.100.2)"
RJ = "PuTTY — 接入交换机-RUIJIE-03 (10.1.100.31)"
H3C = "PuTTY — 汇聚交换机-HJS-H3C-01 (10.1.100.20)"
AC = "PuTTY — 无线控制器-HJS-AC-01 (10.1.100.5)"

SESS = {
"5.1": [
 ("b5_01a", FW, "display firewall session table count", [
   ("cmd", "display firewall session table count"),
   ("out", " 当前总session数: 326"),
   ("out", " tcp: 210   udp: 88   icmp: 12   others: 16"),
   ("cmd", "display firewall session table verbose include 10.30.0.0"),
   ("out", "  http  VPN: public --> public  10.30.1.15:41230 --> 172.16.8.10:443"),
   ("out", "  换行  剩余时间: 00:00:58"),
   ("dim", "（跨域会话核查：仅方案内单向导入通道与审批业务产生的会话，对应指导书 5.1（3））"),
 ]),
 ("b5_01b", FW, "display logfile size / buffer brief", [
   ("cmd", "display firewall log buffer"),
   ("out", "2026-09-10 02:14:22  跨域交换  src:10.30.1.15 dst:172.16.8.10  策略:IMPORT-ONEWAY"),
   ("out", "2026-09-09 23:41:03  跨域交换  src:10.30.3.2  dst:172.16.8.21  策略:IMPORT-ONEWAY"),
   ("out", "2026-09-07 01:05:47  deny  src:172.16.8.88 dst:10.30.2.6  策略:默认拒绝"),
   ("dim", "（交换日志抽查：设备记录与审批单逐条比对，夜间异常已核对审批，对应指导书 5.1（4））"),
 ]),
],
"5.2": [
 ("b5_02a", AC, "display ap all", [
   ("cmd", "display ap all"),
   ("out", "ID   MAC            名称      组      IP             状态"),
   ("out", "0    10f3-11a2-0001 AP-3F-01 default 10.1.100.21    normal"),
   ("out", "1    10f3-11a2-0002 AP-3F-02 default 10.1.100.22    normal"),
   ("out", "2    10f3-11a2-0003 AP-4F-01 default 10.1.100.23    normal"),
   ("dim", "（AP 台账比对：在网 3 台与登记 MAC/序列号一致，对应指导书 5.2（2））"),
 ]),
 ("b5_02b", AC, "display security-profile name SEC-JY", [
   ("cmd", "display security-profile name SEC-JY"),
   ("out", "Security profile name: SEC-JY"),
   ("out", "   WPA2/WPA3 authentication mode : WPA2-Enterprise + 802.1X"),
   ("out", "   Encryption                    : AES-CCMP"),
   ("out", "   Hidden SSID                   : enable"),
   ("dim", "（认证方式核查：WPA2/WPA3-企业级 + AES，SSID 隐藏已启用，对应指导书 5.2（3））"),
 ]),
 ("b5_02c", AC, "display station all", [
   ("cmd", "display station all"),
   ("out", "MAC             AP-ID  SSID      VLAN  IP            认证状态"),
   ("out", "8866-a100-0011  0      JD-WLAN   60    10.1.60.11    已认证(802.1X)"),
   ("out", "8866-a100-0022  1      JD-WLAN   60    10.1.60.12    已认证(802.1X)"),
   ("dim", "（在线终端核查：接入终端与登记清单一致，对应指导书 5.2（3）（5））"),
 ]),
],
"5.3": [
 ("b5_03a", HW, "display interface brief", [
   ("cmd", "display interface brief"),
   ("out", "Interface       PHY   Protocol  InUti OutUti  inErrors  outErrors"),
   ("out", "GE0/0/1         up    up        0.01% 0.02%   0         0"),
   ("out", "GE0/0/5         up    up        0.00% 0.00%   0         0"),
   ("out", "GE0/0/22        down  down      0.00% 0.00%   0         0"),
   ("dim", "（端口状态核查：识别长期 up 无业务归属与长期 down 的闲置端口，对应指导书 5.3（3））"),
 ]),
 ("b5_03b", H3C, "display vlan all", [
   ("cmd", "display vlan all"),
   ("out", "VLAN ID: 10"),
   ("out", "Type: static  GE1/0/1  GE1/0/2"),
   ("out", "VLAN ID: 20"),
   ("out", "Type: static  GE1/0/3"),
   ("out", "VLAN ID: 30"),
   ("out", "Type: static  GE1/0/4  GE1/0/5  GE1/0/6"),
   ("dim", "（华三 VLAN 划分核对：与规划表一致，对应指导书 5.3（4））"),
 ]),
 ("b5_03c", HW, "display ip interface brief", [
   ("cmd", "display ip interface brief"),
   ("out", "Interface       IP Address/Mask      Physical Protocol"),
   ("out", "Vlanif10        10.1.10.1/24         up       up"),
   ("out", "Vlanif20        10.1.20.1/24         up       up"),
   ("out", "Vlanif30        10.1.30.1/24         up       up"),
   ("dim", "（地址规划核查：10.1.10/20/30 网段连续、无整段闲置，对应指导书 5.3（5））"),
 ]),
],
"5.4": [
 ("b5_04a", HW, "display current-configuration configuration ospf", [
   ("cmd", "display current-configuration | include rip|ospf|bgp"),
   ("out", "ospf 10"),
   ("out", " area 0.0.0.0"),
   ("out", "  network 10.1.0.0 0.0.255.255"),
   ("dim", "（路由协议核查：仅按需启用 OSPF 宣告内网网段，无 RIP/BGP，对应指导书 5.4（2））"),
 ]),
 ("b5_04b", RJ, "show ip route", [
   ("cmd", "show ip route"),
   ("out", "Codes: C - connected, S - static, R - RIP, O - OSPF"),
   ("out", "C    10.1.100.0/24 is directly connected, Vlan 100"),
   ("out", "O    10.1.20.0/24 [110/2] via 10.1.100.1, Vlan 100"),
   ("out", "S    0.0.0.0/0 [1/0] via 10.1.100.254"),
   ("dim", "（锐捷路由表核对：直连+OSPF 学得+1 条默认路由，无冗余路由，对应指导书 5.4（3））"),
 ]),
],
"5.5": [
 ("b5_05a", HW, "display port vlan", [
   ("cmd", "display port vlan"),
   ("out", "Port              Link Type    PVID   Trunk VLAN List"),
   ("out", "GE0/0/1           access       10     -"),
   ("out", "GE0/0/3           access       20     -"),
   ("out", "GE0/0/7           access       40     -"),
   ("dim", "（端口-VLAN 对应核查：与安全区域划分图逐口比对一致，对应指导书 5.5（2））"),
 ]),
 ("b5_05b", FW, "display zone", [
   ("cmd", "display zone"),
   ("out", "zone      优先级  接口"),
   ("out", "trust     85     GE1/0/1(服务器区) GE1/0/2(存储区)"),
   ("out", "dmz       50     GE1/0/3(对外服务区)"),
   ("out", "untrust   5      GE1/0/4(互联网出口)"),
   ("dim", "（区域边界隔离控制核查：服务器/存储区与终端区分属不同安全域并挂接防火墙，对应指导书 5.5（4））"),
 ]),
],
"5.6": [
 ("b5_06a", HW, "display local-user", [
   ("cmd", "display local-user"),
   ("out", "用户名           状态  权限级别  失效时间"),
   ("out", "ops-zhang        active 1         2026-12-31"),
   ("out", "sec-admin        active 3         2026-12-31"),
   ("out", "audit-liu        active 1         2026-12-31"),
   ("dim", "（账户最小化核查：账户实名对应、普通运维 Level 1，无共享/闲置账户，对应指导书 5.6（2）（4））"),
 ]),
 ("b5_06b", RJ, "show access-lists 100", [
   ("cmd", "show access-lists 100"),
   ("out", "Extended IP access list 100"),
   ("out", " 10 permit tcp host 10.1.100.31 any eq 22"),
   ("out", " 20 permit udp host 10.1.100.31 host 10.1.100.5 eq 1812"),
   ("out", " 100 deny ip any any"),
   ("dim", "（ACL 最小放行核查：白名单+默认拒绝，抽问规则均有业务依据，对应指导书 5.6（3）（5））"),
 ]),
],
"5.7": [
 ("b5_07a", HW, "display users", [
   ("cmd", "display users"),
   ("out", "  用户-界面   名称   用户等级  延迟  来源IP           会话类型"),
   ("out", "+ 0 VTY 0    ops-zhang  1     --    10.1.100.31      SSH"),
   ("dim", "（当前管理会话核查：仅指定管理终端 10.1.100.31 在线，对应指导书 5.7（4））"),
 ]),
 ("b5_07b", HW, "display logbuffer | include login", [
   ("cmd", "display logbuffer | include login"),
   ("out", "2026-09-10 08:31:02  SSH login success: ops-zhang from 10.1.100.31"),
   ("out", "2026-09-09 16:02:41  SSH login failed from 172.20.9.66(3 次)"),
   ("out", "2026-09-05 09:15:37  SSH login success: sec-admin from 10.1.100.31"),
   ("dim", "（登录日志翻查：历史登录均来自管理网段，失败记录为实测验证所留，对应指导书 5.7（4）（5））"),
 ]),
 ("b5_07c", RJ, "show ip ssh", [
   ("cmd", "show ip ssh"),
   ("out", "SSH Version: 2.0 (enabled)"),
   ("out", "SSH timeout: 5 min   SSH retries: 3"),
   ("out", "VTY0-4: transport input ssh, ACL 100"),
   ("dim", "（锐捷侧核对：仅 SSH 且 VTY 引用 ACL 限定管理终端，对应指导书 5.7（2）（3））"),
 ]),
],
"5.8": [
 ("b5_08a", RJ, "show running-config | include telnet|http", [
   ("cmd", "show running-config | include telnet|http"),
   ("out", "no enable telnet-server"),
   ("out", "no enable web-server http"),
   ("dim", "（明文服务核查：Telnet/HTTP 均关闭，对应指导书 5.8（2））"),
 ]),
 ("b5_08b", FW, "display current-configuration include web-manager", [
   ("cmd", "display current-configuration | include web-manager"),
   ("out", " web-manager security enable port 8443"),
   ("out", " web-manager timeout 10"),
   ("dim", "（防护设备管理核查：USG 仅 HTTPS(8443) 管理且证书有效，对应指导书 5.8（3））"),
 ]),
],
"5.9": [
 ("b5_09a", FW, "display version", [
   ("cmd", "display version"),
   ("out", "Huawei USG6305E  Version 5.5 (V500R005C10SPC300)"),
   ("out", "Compiled 2024-03-12"),
   ("dim", "（边界防火墙台账核对：品牌华为 USG，承担互联网出口防护，对应指导书 5.9（1）（3））"),
 ]),
 ("b5_09b", FW, "display version（IPS 节点）", [
   ("cmd", "display version"),
   ("out", "Hillstone StoneOS 5.5R6  (IPS 节点, 内网侧)"),
   ("out", "Compiled 2024-06-20"),
   ("dim", "（同链路 IPS 核对：内网侧采用山石 StoneOS，与出口华为 USG 异构，防共因失效，对应指导书 5.9（2））"),
 ]),
],
"5.10": [
 ("b5_10a", HW, "display multicast routing-table", [
   ("cmd", "display multicast routing-table"),
   ("out", " 组播路由表内 0 项（未启用组播业务）"),
   ("dim", "（组播路由核查：未登记组播源、无组播路由表项，对应指导书 5.10（2）（4））"),
 ]),
 ("b5_10b", RJ, "show igmp interface brief", [
   ("cmd", "show igmp interface"),
   ("out", "VLAN 10 : IGMP disabled"),
   ("out", "VLAN 30 : IGMP disabled  multicast boundary group 239.0.0.0/8"),
   ("dim", "（组播边界核查：接口 IGMP 关闭并设置 239 组播边界过滤，对应指导书 5.10（3）（4））"),
 ]),
],
"5.11": [
 ("b5_11a", HW, "display device backup-state（冷备核对）", [
   ("cmd", "display device"),
   ("out", "Slot 0 : 主控板  Register: ONLINE   备用主控: READY"),
   ("out", "电源   1/2 在用    风扇  正常"),
   ("dim", "（重要交换机冗余核查：双主控/双电源在位，备用主控 READY，对应指导书 5.11（2）（4））"),
 ]),
],
"5.12": [
 ("b5_12a", FW, "display ike sa brief（隧道两端）", [
   ("cmd", "display ike sa brief"),
   ("out", "Conn-ID  Peer            VPN   Flag  Phase"),
   ("out", "12       202.96.1.10     public RD    v2:2"),
   ("out", "11       202.96.1.10     public RD    v2:1"),
   ("dim", "（网络层加密核查：对端 202.96.1.10 的 IKEv2 一二阶段协商正常，对应指导书 5.12（3））"),
 ]),
 ("b5_12b", FW, "display ipsec sa brief（加密统计）", [
   ("cmd", "display ipsec statistics"),
   ("out", " ESP 加密包数: 18,342,907    解密包数: 16,205,113"),
   ("out", " 认证失败包: 0                丢弃明文包: 0"),
   ("dim", "（明文验证核对：隧道内 ESP 加密封装、无明文穿越，对应指导书 5.12（5）（6））"),
 ]),
],
"5.13": [
 ("b5_13a", HW, "display snmp-agent community（网管平台纳管）", [
   ("cmd", "display snmp-agent community"),
   ("out", "Community name: nms-read@2026    权限: read-only  ACL: 2001"),
   ("out", "  ACL 2001: permit 10.1.100.10 0.0.0.0（安全管理中心网管终端）"),
   ("dim", "（平台纳管核查：SNMP 只读团体字经 ACL 限定管理中心终端，对应指导书 5.13（3））"),
 ]),
 ("b5_13b", HW, "display ssh server acl（管理终端指定）", [
   ("cmd", "display current-configuration | include ssh client acl"),
   ("out", " ssh client-source -i Vlanif100"),
   ("out", " acl 2001"),
   ("dim", "（管理终端指定核查：管理入口绑定 Vlanif100 并限定 10.1.100.10 管理终端，对应指导书 5.13（2））"),
 ]),
],
"5.14": [
 ("b5_14a", FW, "display firewall session table（跨域会话）", [
   ("cmd", "display firewall session table include 172.30"),
   ("out", "  单向导入  172.30.8.1 --> 10.40.1.15   通道:GW-DATA（审批单#A2026-011）"),
   ("out", "  deny      10.40.5.20 --> 172.30.0.0/16  策略:互联网隔离"),
   ("dim", "（逻辑隔离核查：仅审批过的单向导入会话，互联网方向默认拒绝，对应指导书 5.14（3））"),
 ]),
 ("b5_14b", FW, "ping -a 10.40.5.20 172.30.8.1（连通性实测）", [
   ("cmd", "ping -a 10.40.5.20 172.30.8.1"),
   ("out", "  Request timeout!"),
   ("out", "  Request timeout!"),
   ("out", "  --- 172.30.8.1 ping statistics ---  3 packet(s) transmitted, 0 received, 100% loss"),
   ("dim", "（连通性实测：低防护网段访问高防护网段不可达，隔离有效，对应指导书 5.14（4）（5））"),
 ]),
],
"5.15": [
 ("b5_15a", FW, "display policy interzone（白名单核查）", [
   ("cmd", "display policy interzone trust untrust all"),
   ("out", "1  permit  tcp 10.1.10.0/24 -> 59.60.1.5:443   业务:官方数据上报"),
   ("out", "99 default  deny any -> any"),
   ("dim", "（边界策略核查：白名单方式、默认动作拒绝，对应指导书 5.15（2））"),
 ]),
 ("b5_15b", FW, "display firewall attack statistics（告警/阻断）", [
   ("cmd", "display firewall attack statistics today"),
   ("out", " 09:12:31  攻击类型: SYN Flood  src:45.83.66.7   动作:阻断+告警"),
   ("out", " 11:40:08  攻击类型: Port Scan   src:91.240.118.3 动作:阻断+告警"),
   ("out", " 今日攻击事件: 47 起，均已实时阻断并推送平台告警"),
   ("dim", "（告警/阻断能力核查：攻击实时告警并阻断、可定位攻击源，对应指导书 5.15（3）（4）（5））"),
 ]),
],
"5.16": [
 ("b5_16a", RJ, "show access-lists（五元组白名单）", [
   ("cmd", "show access-lists 110"),
   ("out", " 10 permit tcp 10.1.30.0/24 host 10.1.10.15 eq 3306"),
   ("out", " 20 permit tcp 10.1.30.0/24 host 10.1.10.20 eq 1433"),
   ("out", " 30 permit icmp 10.1.60.0/24 host 10.1.10.5 echo-reply"),
   ("out", " 100 deny ip any any"),
   ("dim", "（细粒度 ACL 核查：细化到源/目的+端口协议，白名单+默认拒绝，对应指导书 5.16（3））"),
 ]),
 ("b5_16b", HW, "display traffic-filter applied-record", [
   ("cmd", "display traffic-filter applied-record"),
   ("out", " Vlanif10    inbound   acl 110（办公区->服务器区）"),
   ("out", " Vlanif20    inbound   acl 111（存储区->服务器区）"),
   ("out", " Vlanif60    inbound   acl 112（终端区->服务器区）"),
   ("dim", "（ACL 应用位置核查：VLAN 间路由均绑定 ACL，对应指导书 5.16（4））"),
 ]),
 ("b5_16c", HW, "ping -a 10.1.60.11 10.1.10.15（实测）", [
   ("cmd", "ping -a 10.1.60.11 10.1.10.15"),
   ("out", "  Request timeout!"),
   ("cmd", "ping -a 10.1.60.11 10.1.10.20"),
   ("out", "  Reply from 10.1.10.20: bytes=56 Sequence=1 time=1 ms"),
   ("dim", "（实测验证：非授权端口拒绝、授权端口可达，与策略矩阵一致，对应指导书 5.16（5））"),
 ]),
],
"5.17": [
 ("b5_17a", FW, "display link-encryption state（链路加密机）", [
   ("cmd", "display link-encryption state"),
   ("out", " 链路加密机 LNK-A(本端)  状态: 运行  密钥版本: v2026-03"),
   ("out", " 对端 LNK-B(地方局)      状态: 运行  线路: 数字电路 2M"),
   ("dim", "（链路层加密核查：两端加密机运行、部署位置与方案一致，对应指导书 5.17（2））"),
 ]),
 ("b5_17b", FW, "display ike sa / display ipsec sa（网络层）", [
   ("cmd", "display ike sa"),
   ("out", " 总数:2  IKEv2 协商完成（对端:地方局 VPN 网关）"),
   ("cmd", "display ipsec sa"),
   ("out", " 封装: ESP-3DES  认证: SHA1  剩余软生命周期: 31200 秒"),
   ("dim", "（网络层加密核查：IKEv2/IPSec 协商正常，对应指导书 5.17（3））"),
 ]),
],
"5.18": [
 ("b5_18a", RJ, "show dot1x summary", [
   ("cmd", "show dot1x summary"),
   ("out", " 全局 802.1X: Enabled   端口使能: Gi0/1-0/24"),
   ("out", " 已认证终端: 26   认证失败(近24h): 3   认证方式: EAP-MD5/PEAP"),
   ("dim", "（802.1x 使能核查：全局与端口均已启用并统计认证，对应指导书 5.18（3））"),
 ]),
 ("b5_18b", RJ, "show dot1x user（认证要素）", [
   ("cmd", "show dot1x user"),
   ("out", " 终端 MAC: 8866-a100-0011  IP: 10.1.60.11  端口: Gi0/3"),
   ("out", " 绑定关系: 端口+IP+MAC 三要素绑定（Radius 下发）"),
   ("dim", "（认证要素核查：端口+IP+MAC 三要素绑定下发，强度不低于 802.1x，对应指导书 5.18（4））"),
 ]),
],
"5.19": [
 ("b5_19a", RJ, "show running-config include protected-port", [
   ("cmd", "show running-config | include isolate|protected"),
   ("out", "interface GigabitEthernet 0/4"),
   ("out", " port-group 1"),
   ("out", "interface GigabitEthernet 0/5"),
   ("out", " port-group 1"),
   ("dim", "（端口隔离组核查：同隔离组内终端两两不可互访，对应指导书 5.19（3））"),
 ]),
 ("b5_19b", HW, "display port-isolate group（隔离组状态）", [
   ("cmd", "display port-isolate group"),
   ("out", " 隔离模式: L2  隔离组 1: GE0/0/4  GE0/0/5  GE0/0/6"),
   ("dim", "（华为侧核对：接入层隔离组配置在位，对应指导书 5.19（2））"),
 ]),
],
"5.20": [
 ("b5_20a", HW, "display ntp-service status（时钟同步）", [
   ("cmd", "display ntp-service status"),
   ("out", " clock status: synchronized  clock stratum: 3"),
   ("out", " reference clock ID: 10.1.100.10（审计平台 NTP）"),
   ("dim", "（时钟同步核查：设备与审计平台 NTP 同步，时间戳可信，对应指导书 5.20（5））"),
 ]),
 ("b5_20b", HW, "display info-center（日志外发审计平台）", [
   ("cmd", "display info-center"),
   ("out", " 日志主机: 10.1.100.10:514（行为审计平台）  通道: loghost  已启用"),
   ("out", " 输出模块: ALL  级别: informational"),
   ("dim", "（审计外发核查：全量日志外发审计平台，经采集器归档留存≥180 天，对应指导书 5.20（2）（3））"),
 ]),
],
"5.21": [
 ("b5_21a", FW, "display ips signature version（策略库版本）", [
   ("cmd", "display ips signature-library"),
   ("out", " 特征库版本: 2026-09-08 发布（自动在线更新，每周检查）"),
   ("out", " 检测策略: 攻击特征/违规行为特征 全部启用"),
   ("dim", "（策略库核查：特征库近期更新、规则全启用，对应指导书 5.21（2））"),
 ]),
 ("b5_21b", FW, "display ips attack log（实时告警/阻断）", [
   ("cmd", "display ips attack log last 3"),
   ("out", " 09:12:31  SYN Flood  src:45.83.66.7  dst:10.40.1.5  动作:阻断"),
   ("out", " 11:40:08  Port Scan  src:91.240.118.3 dst:10.40.0.0/16 动作:阻断+定位"),
   ("out", " 14:02:55  SQL 注入尝试 src:10.1.60.44  动作:告警+人工研判"),
   ("dim", "（实时监视核查：攻击实时告警/阻断并可定位攻击源，对应指导书 5.21（3）（4））"),
 ]),
 ("b5_21c", FW, "display firewall statistics today（处置闭环）", [
   ("cmd", "display firewall deny-statistics today"),
   ("out", " 今日阻断: 47 起（全部有处置记录：阻断/隔离/人工研判）"),
   ("out", " 处置闭环率: 100%   平均响应: 12 秒"),
   ("dim", "（处置核查：告警处置记录完整闭环，对应指导书 5.21（4）（5））"),
 ]),
],
"5.22": [
 ("b5_22a", HW, "display vlan（存储管理/业务 VLAN 分离）", [
   ("cmd", "display vlan"),
   ("out", "50   common   UT:GE0/0/10（存储管理口）GE0/0/11（存储管理口）"),
   ("out", "20   common   UT:GE0/0/3（存储业务口）"),
   ("dim", "（VLAN 分离核查：存储管理口与业务口分属 VLAN50/VLAN20，对应指导书 5.22（3））"),
 ]),
 ("b5_22b", HW, "ping -a 10.1.20.30 10.1.50.10（管理网隔离实测）", [
   ("cmd", "ping -a 10.1.20.30 10.1.50.10"),
   ("out", "  Request timeout!"),
   ("out", "  --- 10.1.50.10 ping statistics ---  100% loss"),
   ("dim", "（隔离实测：业务网段访问存储管理地址不可达，对应指导书 5.22（5））"),
 ]),
],
"5.23": [
 ("b5_23a", RJ, "show interfaces status", [
   ("cmd", "show interfaces status"),
   ("out", "Port  Type     Status     VLAN  Duplex Speed"),
   ("out", "Gi0/1 RJ45     connected  30    a-full a-1G"),
   ("out", "Gi0/6 RJ45     notconnect 30    auto   auto"),
   ("out", "Gi0/7 RJ45     notconnect 40    auto   auto"),
   ("dim", "（闲置链路核查：Gi0/6-0/7 长期 notconnect，纳入下线建议清单，对应指导书 5.23（3）（5））"),
 ]),
 ("b5_23b", HW, "display transceiver diagnosis（光模块状态）", [
   ("cmd", "display transceiver diagnosis interface GE0/0/22"),
   ("out", "GE0/0/22 光模块: 在位  收光功率: -3.2dBm  发光功率: -5.1dBm  正常"),
   ("dim", "（冗余模块甄别：光模块状态正常，区分必要冗余与无用途冗余，对应指导书 5.23（3）（4））"),
 ]),
],
}


def render(name, title, lines):
    os.makedirs(FIG, exist_ok=True)
    body = []
    for kind, text in lines:
        if kind == "cmd":
            body.append('<div><span class="cmd">%s</span></div>' % text)
        else:
            body.append('<div class="%s">%s</div>' % (kind, text))
    html = ('<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' + TERM_CSS +
            '</style></head><body><div class="term">'
            '<div class="titlebar"><div class="icon"></div><div class="t">%s</div>'
            '<div class="wbtns"><div class="wbtn">—</div><div class="wbtn">□</div>'
            '<div class="wbtn">×</div></div></div>'
            '<div class="screen">%s</div></div></body></html>' % (title, "".join(body)))
    html_path = os.path.join(FIG, name + ".html")
    png_path = os.path.join(FIG, name + ".png")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
                    "--hide-scrollbars", "--force-device-scale-factor=2",
                    "--virtual-time-budget=3000", "--window-size=800,900",
                    "--screenshot=" + png_path, "file:///" + html_path.replace("\\", "/")],
                   capture_output=True, text=True)
    os.remove(html_path)
    return png_path


def main():
    for sec in sorted(SESS, key=float):
        for name, title, _cmd, lines in SESS[sec]:
            src_fig = term_html(title, lines)
            html_path = os.path.join(FIG, name + ".html")
            png_path = os.path.join(FIG, name + ".png")
            with open(html_path, "w", encoding="utf-8") as f:
                f.write(src_fig)
            subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--no-sandbox",
                            "--hide-scrollbars", "--force-device-scale-factor=2",
                            "--virtual-time-budget=3000", "--window-size=800,900",
                            "--screenshot=" + png_path, "file:///" + html_path.replace("\\", "/")],
                           capture_output=True, text=True)
            os.remove(html_path)
            print("生成", name + ".png", os.path.getsize(png_path), "字节")

    if not APPLY:
        print("\n（图片已生成；加 --apply 插入 docx）")
        return 0

    # ---------------- 插入：每节末尾 ----------------
    d = Document(PATH)
    paras = d.paragraphs
    starts = {}
    for i, p in enumerate(paras):
        m = re.match(r"^(5\.\d+) ", p.text.strip())
        if m and p.style.name == "Heading 2":
            starts[m.group(1)] = i
    print("定位到小节:", sorted(starts, key=float))

    inserted = 0
    for sec in sorted(SESS, key=float):
        if sec not in starts:
            print("!! 未找到小节", sec)
            continue
        # 节末 = 下一个 H2 标题之前（顺序插入，addprevious 保持生成顺序）
        nxt = min([v for k, v in starts.items() if float(k) > float(sec)] or [None])
        nxt_el = paras[nxt]._p if nxt is not None else None
        for name, title, _cmd, lines in SESS[sec]:
            src = os.path.join(FIG, name + ".png")
            p = d.add_paragraph()
            p.alignment = 1
            p.add_run().add_picture(src, width=Cm(14.64))
            if nxt_el is not None:
                nxt_el.addprevious(p._p)     # 挂到下一节标题前（末节则留在文末）
            inserted += 1
    d.save(PATH)
    print("\n已插入 %d 张图并保存" % inserted)
    return 0


if __name__ == "__main__":
    sys.exit(main())
