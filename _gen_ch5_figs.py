# -*- coding: utf-8 -*-
"""第5章网络安全示意图批量生成：Mermaid → Chrome headless → PIL 裁边 → PNG。"""
import http.server
import threading
import subprocess
import os
import time
import tempfile

BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "_ch5_figs")
os.makedirs(FIG, exist_ok=True)

THEME_INIT = """mermaid.initialize({
  startOnLoad: true,
  theme: 'base',
  themeVariables: {
    primaryColor: '#eef3fa', primaryBorderColor: '#1a3c6e', primaryTextColor: '#1f2937',
    lineColor: '#475569', fontSize: '15px',
    clusterBkg: '#f7fafc', clusterBorder: '#94a3b8'
  },
  flowchart: { curve: 'basis', nodeSpacing: 34, rankSpacing: 44 }
});"""

FIGS = {
"fig5_1": ("flowchart TB\n"
           "  A[流程与文档类\\n5.1-5.2] --- B[架构与区域类\\n5.3-5.5] --- C[设备配置类\\n5.6-5.8] "
           "--- D[链路与冗余类\\n5.9-5.12/5.17] --- E[隔离与审计类\\n5.13-5.16/5.18-5.23]\n"),
"fig5_2": ("flowchart TD\n"
           "  A[提出跨网跨域交换申请] --> B{网络管理部门与主管保密部门审批}\n"
           "  B -- 通过 --> C[安全保密方案设计]\n  B -- 不通过 --> H[终止或整改]\n"
           "  C --> D{专家评审}\n  D -- 不通过 --> C\n"
           "  D -- 通过 --> E[按方案实施交换]\n  E --> F[交换日志留存]\n"
           "  F --> G[定期比对：交换行为与审批范围]\n"),
"fig5_3": ("flowchart TB\n"
           "  CORE((核心交换机))\n"
           "  subgraph Z[安全区域划分]\n"
           "    SRV[服务器区 VLAN10]\n    STG[数据存储区 VLAN20]\n    OFF[办公终端区 VLAN30]\n    MGT[安全管理区 VLAN40]\n  end\n"
           "  CORE --> SRV & STG & OFF & MGT\n"
           "  OFF -. 须经集中服务中转 .-> SRV\n  OFF -. 管理通道 .-> MGT\n"),
"fig5_4": ("flowchart LR\n"
           "  ADM[指定管理终端<br/>仅 ACL 放行网段]\n  SW((网络设备))\n  FW((安全防护设备))\n"
           "  ADM ==>|SSH 2.0 加密| SW\n  ADM ==>|HTTPS 加密| FW\n"
           "  T1[Telnet 明文管理] -. 禁止 .-> SW\n  T2[HTTP 明文管理] -. 禁止 .-> FW\n"
           "  T3[非管理终端] -. ACL 拒绝 .-> SW\n"),
"fig5_5": ("flowchart LR\n"
           "  INET[互联网] ===|物理隔离| IN[内部网络]\n"
           "  NJD[非JD网络] ===|物理隔离| IN\n  L1[一级防护网络] ===|物理隔离| IN\n"
           "  subgraph LG[二/三/四级网络 逻辑隔离]\n"
           "    IN --- FW((边界防火墙))\n    FW --- L2[二级防护网络]\n    FW --- L3[三级防护网络]\n    FW --- L4[四级防护网络]\n  end\n"
           "  FW --- IPS[IPS/网闸/单向传输<br/>告警·审计·阻断·定位]\n"),
"fig5_6": ("sequenceDiagram\n"
           "  participant T as 接入终端<br/>(端口+IP+MAC)\n"
           "  participant S as 接入交换机<br/>(802.1x 使能)\n"
           "  participant R as 认证服务器<br/>(Radius/准入)\n"
           "  T->>S: 发起网络接入\n"
           "  S->>T: 要求 802.1x 认证\n"
           "  T->>S: 提交认证凭据\n"
           "  S->>R: 转发认证请求（绑定端口/IP/MAC）\n"
           "  alt 认证通过\n    R-->>S: 授权（放行业务 VLAN）\n    S-->>T: 接入业务网络\n  else 认证失败/未注册\n    R-->>S: 拒绝\n    S-->>T: 拒绝接入或进入隔离区\n  end\n"),
"fig5_7": ("flowchart TB\n"
           "  subgraph Z[各安全区域与边界出口]\n"
           "    Z1[服务器区] --- Z2[办公终端区] --- Z3[边界出口]\n  end\n"
           "  Z1 --> AUD[全网行为审计系统<br/>日志留存 ≥ 180 天]\n"
           "  Z3 --> IDS[IDS/IPS/态势感知<br/>实时告警·阻断·定位]\n"
           "  AUD --> MC((网络安全管理中心))\n  IDS --> MC\n"
           "  MC -. 处置闭环 .-> Z3\n"),
"fig5_8": ("flowchart TB\n"
           "  STG((数据存储设备))\n"
           "  STG -->|管理口| MV[管理网络 VLAN40]\n"
           "  STG -->|业务口| BV[业务应用网络 VLAN20]\n"
           "  BV -. 不互通 .-> MV\n"),
}

for name, mmd in FIGS.items():
    html = ('<!DOCTYPE html><html><head><meta charset="UTF-8">'
            '<style>body{margin:0;background:#fff;font-family:"Microsoft YaHei",sans-serif;}</style>'
            '<script src="mermaid.min.js"></script></head><body>'
            '<pre class="mermaid">' + mmd + '</pre>'
            '<script>' + THEME_INIT + '</script></body></html>')
    open(os.path.join(FIG, name + '.html'), 'w', encoding='utf-8').write(html)

srv = http.server.HTTPServer(('127.0.0.1', 18963), http.server.SimpleHTTPRequestHandler)
threading.Thread(target=srv.serve_forever, daemon=True).start()
time.sleep(0.4)
chrome = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
prof = tempfile.mkdtemp(prefix='mmd_render_')
for name in FIGS:
    out_png = os.path.join(FIG, name + '.png')
    r = subprocess.run([chrome, '--headless=new', '--disable-gpu', f'--user-data-dir={prof}_{name}',
                        '--hide-scrollbars', '--force-device-scale-factor=2',
                        '--window-size=1700,1400', f'--screenshot={out_png}',
                        '--virtual-time-budget=9000',
                        f'http://127.0.0.1:18963/_ch5_figs/{name}.html'],
                       capture_output=True, text=True, timeout=90)
    print(name, 'size=', os.path.getsize(out_png) if os.path.exists(out_png) else 0)
srv.shutdown()

from PIL import Image, ImageChops
for name in FIGS:
    path = os.path.join(FIG, name + '.png')
    img = Image.open(path).convert('RGB')
    bg = Image.new('RGB', img.size, (255, 255, 255))
    bbox = ImageChops.difference(img, bg).getbbox()
    if bbox:
        pad = 16
        bbox = (max(0, bbox[0]-pad), max(0, bbox[1]-pad),
                min(img.width, bbox[2]+pad), min(img.height, bbox[3]+pad))
        img.crop(bbox).save(path)
    im2 = Image.open(path)
    w, h = im2.size
    ratio = w / h
    warn = ' ⚠比例需注意' if (ratio < 0.35 or ratio > 6) else ''
    print(name, f'裁边后 {w}×{h} (比例{ratio:.1f}:1){warn}')
print('全部完成')
