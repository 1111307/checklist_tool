# -*- coding: utf-8 -*-
"""按指导书内容生成核查表对象矩阵（136 项 × 11 列）v2。

规则（用户定义）：指导书某个检查项小节写明了哪个对象的核查方法就勾 √，没写就横杠 —。
信号源：
  1. 子节标题（H3/H4/H5，去编号）含对象名，如 4.9.1Win7、WinXP / 1.7.1 Mysql / 3.1.3 通用核查方法
  2. 正文行内平台前缀：「Win7/WinXP：」「麒麟：」「Nginx：」「Tomcat：」（通用核查方法里最常用）
  3. 「XX操作系统下的核查方法」伪标题、正文单独成行的对象名
  4. 正文中具体产品名（MySQL / SQL Server / 达梦 / Redis / Nginx / Tomcat 出现即算）
  5. 网络设备：第 5 章整章强制 √，其余章节按强信号（display/show 设备命令、华为/华三/锐捷、
     USG/VRP、网闸、交换机、防火墙设备、网络设备）
第 6-9 章与 10.1 为现场/文档/协议类，不含上述对象，输出全横杠。

输出：_matrix_from_guide.json（136 项覆盖集），供核查表更新与标注/报告联动。
"""
import re, json
from docx import Document

doc = Document('配置核查作业指导书_v2.2.docx')

# 列顺序（与核查表 C~M 一致，华为交换机插在 Tomcat 之后）
COLS = ['Win7', 'WinXP', '中标麒麟', '银河麒麟', 'Nginx', 'Tomcat', '华为交换机', 'Mysql', 'SQLSever', '达梦', 'redis']

# 正文行内平台前缀（通用核查方法里的「麒麟：cmd」「Win7/WinXP：cmd」写法）
INLINE_PLAT = [
    ('Win7',    re.compile(r'(^|[，,、；;\s（(])Win7\s*[/、,，]\s*WinXP\s*[：:]')),
    ('WinXP',   re.compile(r'(^|[，,、；;\s（(])Win7\s*[/、,，]\s*WinXP\s*[：:]')),
    ('Win7',    re.compile(r'(^|[，,、；;\s（(])Win7\s*[：:]')),
    ('WinXP',   re.compile(r'(^|[，,、；;\s（(])WinXP\s*[：:]')),
    ('中标麒麟', re.compile(r'(^|[，,、；;\s（(])中标麒麟\s*[：:]')),
    ('银河麒麟', re.compile(r'(^|[，,、；;\s（(])银河麒麟\s*[：:]')),
    ('中标麒麟', re.compile(r'(^|[，,、；;\s（(])麒麟\s*[：:]')),
    ('银河麒麟', re.compile(r'(^|[，,、；;\s（(])麒麟\s*[：:]')),
    # 裸称「Windows：」（如 5.19「（Windows：防火墙入站默认阻止；麒麟：…）」）——两端并列的平台前缀写法
    ('Win7',    re.compile(r'(^|[，,、；;\s（(])Windows\s*[：:]', re.I)),
    ('WinXP',   re.compile(r'(^|[，,、；;\s（(])Windows\s*[：:]', re.I)),
    ('Nginx',   re.compile(r'(^|[，,、；;\s（(])Nginx\s*[：:]')),
    ('Tomcat',  re.compile(r'(^|[，,、；;\s（(])Tomcat\s*[：:]')),
]
# 正文/标题中出现对象名即算覆盖
PROD_RX = [
    ('Win7', re.compile(r'Win7|Windows\s?7|Windows7/WindowXP|Windows操作系统|Windows 操作系统', re.I)),
    ('WinXP', re.compile(r'WinXP|Windows\s?XP|Windows7/WindowXP|Windows操作系统|Windows 操作系统|Win7/XP|Win7、WinXP', re.I)),
    ('中标麒麟', re.compile(r'中标麒麟', re.I)),
    ('银河麒麟', re.compile(r'银河麒麟', re.I)),
    ('Nginx', re.compile(r'\bNginx\b', re.I)),
    ('Tomcat', re.compile(r'\bTomcat\b', re.I)),
    ('Mysql', re.compile(r'\bMySQL\b|\bMysql\b', re.I)),
    ('SQLSever', re.compile(r'SQL\s?Server|SQLServer|MSSQL', re.I)),
    ('达梦', re.compile(r'达梦|DM8|disql|dmserver', re.I)),
    ('redis', re.compile(r'\bRedis\b|redis-cli', re.I)),
]
# 网络设备强信号
NET_RX = re.compile(
    r'\bdisplay \w|'
    r'\bshow (running-config|version|vlan|ip ssh|users|access-lists|dot1x|interfaces status|session|ip route|configuration|logbuffer|acl|mac-address|lldp)\b|'
    r'华为|华三|锐捷|USG|VRP|网闸|交换机|防火墙设备|登录\S*防火墙|网络边界设备|网络设备', re.I)
# 正文单独成行的对象名
MARK_RX = re.compile(
    r'^(Windows7/WindowXP|Win7\\WinXP|Win7\s*[/、]\s*WinXP|Windows\s?7|Windows\s?XP|中标麒麟|银河麒麟|中标麒麟、银河麒麟|'
    r'MySQL|Mysql|SQL Server|SQLServer|达梦|Redis|Nginx|Tomcat|Apache)(\s|$|：|:)', re.I)


def strip_num(t):
    return re.sub(r'^\s*\d+(\.\d+)+\s*', '', t).strip()


# ---- 遍历文档：按 H2 切分检查项小节 ----
items = []
cur_ch = None
cur = None
body = []
for p in doc.paragraphs:
    st = p.style.name
    txt = p.text.strip()
    if st == 'Heading 1':
        if cur:
            items.append((cur_ch, cur, '\n'.join(body)))
        cur_ch = txt
        cur = None
        body = []
    elif st == 'Heading 2':
        if cur:
            items.append((cur_ch, cur, '\n'.join(body)))
        m = re.match(r'(\d+\.\d+)\s+(.*)', txt)
        cur = (m.group(1), m.group(2)) if m else (txt, txt)
        body = []
    else:
        if cur and txt:
            body.append(f'[{st}] {txt}')
if cur:
    items.append((cur_ch, cur, '\n'.join(body)))

# ---- 对象覆盖检测 ----
matrix = {}
for ch, (code, title), body_txt in items:
    bh = int(code.split('.')[0])
    hits = set()
    if bh == 5:
        hits.add('华为交换机')          # 第 5 章整章为网络设备专题
    for line in body_txt.split('\n'):
        if '] ' not in line:
            continue
        tag, rest = line.split('] ', 1)
        st = tag[1:]
        t = strip_num(rest) if st in ('Heading 3', 'Heading 4', 'Heading 5') else rest
        for obj, rx in PROD_RX:
            if rx.search(t):
                hits.add(obj)
        for obj, rx in INLINE_PLAT:
            if rx.search(t):
                hits.add(obj)
        if bh not in (6, 7, 8, 9, 10) and NET_RX.search(t):
            hits.add('华为交换机')
    matrix[code] = {'ch': ch, 'title': title, 'objs': [c for c in COLS if c in hits]}

json.dump(matrix, open('_matrix_from_guide.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('生成 _matrix_from_guide.json，共', len(matrix), '项')

print(f'\n{"编号":6s} {"检查项(截断)":26s} ' + ' '.join(f'{c[:2]:>3s}' for c in COLS))
for code in sorted(matrix, key=lambda x: (int(x.split('.')[0]), int(x.split('.')[1]))):
    m = matrix[code]
    row = ' '.join('√' if c in m['objs'] else '—' for c in COLS)
    print(f'{code:6s} {m["title"][:26]:26s} {row}')
