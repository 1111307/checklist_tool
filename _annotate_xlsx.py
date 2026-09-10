# -*- coding: utf-8 -*-
"""核查表第13列「工具自动验证」详情标注：脚本名 + 可自动部分 + 需人工部分。
分类严格对齐交叉验证报告结论（86 已实现/50 部分实现/0 未实现），
详情取自 15 个脚本源码的实际判定分支。用法：python _annotate_xlsx.py [--apply]
"""
import re, glob, os, sys, math
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Alignment

XLSX = '配置核查表_v2.0.0_标注自动验证.xlsx'
APPLY = '--apply' in sys.argv

# ---------- 1. 脚本覆盖提取 ----------
def extract_sh(path):
    data = open(path, encoding='utf-8').read()
    d = defaultdict(set)
    for m in re.finditer(r'add_result\s+"([^"]+)"\s+"[^"]*"\s+"[^"]*"\s+"([^"]+)"', data):
        d[m.group(1)].add(m.group(2))
    return d
def extract_vbs(path):
    data = open(path, encoding='gbk').read()
    data = re.sub(r'_\s*\r?\n\s*', '', data)
    d = defaultdict(set)
    for m in re.finditer(r'AddResult\s*\(?\s*"([^"]+)"\s*,\s*"[^"]*"\s*,\s*"[^"]*"\s*,\s*"([^"]+)"', data):
        d[m.group(1)].add(m.group(2))
    return d

scripts = {}
for f in sorted(glob.glob('win/check_*.vbs')):
    scripts['win/' + os.path.basename(f).replace('.vbs', '')] = extract_vbs(f)
for f in sorted(glob.glob('kylin/check_*.sh')):
    scripts['kylin/' + os.path.basename(f).replace('.sh', '')] = extract_sh(f)

def expand(code):
    out = set()
    m = re.match(r'^(\d+\.\d+)-(\d+)$', code)
    if m:
        pre = m.group(1)
        out.add(pre); out.add(pre.split('.')[0] + '.' + m.group(2))
        return out
    m = re.match(r'^(\d+\.\d+)_', code)
    if m:
        out.add(m.group(1)); return out
    out.add(code)
    return out

cov = defaultdict(lambda: defaultdict(set))
for sname, d in scripts.items():
    for code, sts in d.items():
        for c in expand(code):
            cov[c][sname] |= sts

# ---------- 2. 报告结论 ----------
md = open('测评报告/指导书与核查工具交叉验证报告.md', encoding='utf-8').read()
report = {}
for m in re.finditer(r'^\| (\d+\.\d+) \|[^|]+\|[^|]+\| (已实现|部分实现|未实现) \|', md, re.M):
    report[m.group(1)] = m.group(2)

# ---------- 3. 手工文案（脚本详情/人工边界已逐项核对源码） ----------
HAND = {
    # 第1-4章 部分实现 15 项
    '1.16': '部分可自动化（脚本无法判定应用层参数校验逻辑，需人工或 AWVS 类工具核查）',
    '3.4':  '部分可自动化（载体销毁为物理过程，脚本不适用；需人工核查销毁档案与影像记录）',
    '3.12': '部分可自动化（脚本可查采集/同步服务进程，采集范围与业务需求比对需人工）',
    '3.14': '部分可自动化（OS 层权限与审计可查，大数据平台统一管控需平台界面人工核查）',
    '4.1':  '部分可自动化（OS/数据库备份任务可查，应用自身备份功能需登录应用核查）',
    '4.9':  '部分可自动化（curl 可查安全响应头，SQL注入/XSS防护能力需 WAF/代码审计）',
    '4.10': '部分可自动化（上传白名单等配置可查，执行代码校验需 nikto 扫描或人工）',
    '4.13': '部分可自动化（登录来源 IP 可比对，管理终端专设专用需现场核对台账）',
    '4.15': '部分可自动化（证书工具安装状态可查，签名验证/密级标识功能需人工核查）',
    '4.20': '部分可自动化（国密 SM2/3/4 配置可 grep，自主设计认定需人工）',
    '4.21': '部分可自动化（多因素与数字证书认证需人工核查 Windows Hello/国密证书库）',
    '4.26': '部分可自动化（代码级漏洞挖掘需渗透测试/代码审计工具，脚本不适用）',
    '4.27': '部分可自动化（php upload_max_filesize 等校验配置可查，校验有效性需人工测试）',
    '4.28': '部分可自动化（curl 可查安全头，四类攻击防御能力需 WAF/渗透测试验证）',
    '4.29': '部分可自动化（统一权限管理平台需人工核查 IAM/SSO 接入与变更记录）',
    # 特定已实现项的准确表述
    '2.3':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查自动登录/空口令账户限制）',
    '2.7':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查系统盘文件保护状态）',
    '3.6':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 DLP/终端管控客户端部署）',
    '3.1':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 OS 层磁盘加密；数据库 TDE 需人工）',
    '10.1': '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 TLS/SSH 协议版本）',
    # 第5章 部分（脚本提取信号 + 台账）
    '5.9':  '部分可自动化（用 check_network.sh 提取设备型号，异构部署结论需比对台账）',
    '5.22': '部分可自动化（用 check_network.sh 提取存储网/业务网划分信号，结论需比对台账）',
    '5.23': '部分可自动化（用 check_network.sh 计算端口闲置率，配备必要性结论需人工评估）',
}
PROD_ORDER = ['MySQL', 'Redis', '达梦', 'SQL Server', 'Nginx', 'Tomcat']
PROD_KEY = {'mysql': 'MySQL', 'redis': 'Redis', 'dm': '达梦',
            'sqlserver': 'SQL Server', 'nginx': 'Nginx', 'tomcat': 'Tomcat'}

def compose(code):
    ch = int(code.split('.')[0])
    if code in HAND:
        return HAND[code]
    rep = report.get(code)
    if rep is None:
        return None
    if ch in (6, 7, 8, 9):
        return '需人工（现场/文档/台账类核查）'
    if ch == 5:
        if rep == '已实现':
            return '可自动化（用 check_network.sh 采集设备回显后自动判定）'
        return '需人工（台账/平台核查）'
    # 第1-4章
    if rep == '部分实现':
        return None  # 应全部在 HAND 中
    auto = [s for s, sts in cov.get(code, {}).items() if 'pass' in sts or 'fail' in sts]
    os_win = 'win/check_xp7' in auto
    os_ky = 'kylin/check_kylin' in auto
    comps = {PROD_KEY[s.split('/')[1].replace('check_', '')]
             for s in auto
             if s.split('/')[1] not in ('check_xp7', 'check_kylin', 'check_network')}
    prods = '、'.join(p for p in PROD_ORDER if p in comps)
    if os_win and os_ky:
        return f'可自动化（用 check_xp7.vbs、check_kylin.sh{" 及 " + prods + " 组件脚本" if prods else ""}自动核查）'
    if os_win:
        return f'可自动化（用 check_xp7.vbs{" 及 " + prods + " 组件脚本" if prods else ""}自动核查；麒麟 OS 层需人工确认）'
    if os_ky:
        return f'可自动化（用 check_kylin.sh{" 及 " + prods + " 组件脚本" if prods else ""}自动核查；Windows OS 层需人工确认）'
    if prods:
        return f'可自动化（用 {prods} 组件脚本自动核查）'
    return None

# ---------- 4. 行号映射 ----------
wb = load_workbook(XLSX)
ws = wb.active
chapters = {'系统安全': 1, '用户安全': 2, '数据安全': 3, '应用安全': 4, '网络安全': 5, '物理安全': 6}
rowmap = {}
cur = None; seq = 0
for r in range(4, ws.max_row + 1):
    c1 = ws.cell(row=r, column=1).value
    if c1 and str(c1).strip() in chapters:
        cur = chapters[str(c1).strip()]; seq = 0
    seq += 1
    if cur == 6:
        if seq <= 9: code = f'6.{seq}'
        elif seq <= 14: code = f'7.{seq-9}'
        elif seq <= 18: code = f'8.{seq-14}'
        elif seq <= 22: code = f'9.{seq-18}'
        else: code = '10.1'
    else:
        code = f'{cur}.{seq}'
    rowmap[r] = code

# R139 为备注行（10.1 之后），不标注
note_rows = [r for r, c in rowmap.items() if c == '10.1' and r > 138]

# ---------- 5. 生成 + 一致性校验 ----------
CH5_MANUAL = {'5.1', '5.2', '5.12', '5.13', '5.14', '5.15', '5.17', '5.20', '5.21'}
errors, texts = [], {}
from collections import Counter
dist = Counter()
for r in range(4, 139):
    code = rowmap[r]
    if r in note_rows:
        continue
    txt = compose(code)
    if txt is None:
        errors.append(f'R{r} {code}: 无文案（compose 返回空）')
        continue
    rep = report.get(code)
    # 分类与报告一致性
    if txt.startswith('可自动化') and rep != '已实现':
        errors.append(f'R{r} {code}: 标可自动化但报告={rep}')
    if txt.startswith('部分可自动化') and rep != '部分实现':
        errors.append(f'R{r} {code}: 标部分但报告={rep}')
    if txt.startswith('需人工'):
        if code in CH5_MANUAL:
            pass
        elif int(code.split('.')[0]) in (6, 7, 8, 9):
            if rep != '部分实现':
                errors.append(f'R{r} {code}: 标需人工但报告={rep}')
        else:
            errors.append(f'R{r} {code}: 非网络/6-9章标需人工')
    texts[r] = txt
    cat = '可自动化' if txt.startswith('可自动化') else ('部分可自动化' if txt.startswith('部分') else '需人工')
    dist[cat] += 1

if errors:
    print('!! 校验失败：')
    for e in errors:
        print('  ', e)
    sys.exit(1)

print(f'=== 干跑通过：{len(texts)} 条文案，分布 {dict(dist)} ===')
print(f'预期：可自动化 86 / 部分可自动化 18 / 需人工 31（备注行 1 条不标注）')
if not (dist['可自动化'] == 86 and dist['部分可自动化'] == 18 and dist['需人工'] == 31):
    print('!! 分布与预期不符，中止')
    sys.exit(1)
if not APPLY:
    print('（干跑模式，未写盘。加 --apply 应用）')
    sys.exit(0)

# ---------- 6. 写入 ----------
for r, txt in texts.items():
    ws.cell(row=r, column=13).value = txt

# M 列加宽 + 自动换行；行高按需抬高（只升不降，自动行高行不动）
ws.column_dimensions['M'].width = 55
for r, txt in texts.items():
    c = ws.cell(row=r, column=13)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    lines = math.ceil(len(txt) / 26)  # 宽 55 ≈ 每行 26 个汉字
    need = {1: 17, 2: 34, 3: 50}.get(lines, 50)
    h = ws.row_dimensions[r].height
    if h is not None and h < need:
        ws.row_dimensions[r].height = need

wb.save(XLSX)
print(f'已写入 {XLSX}（M 列宽 55、自动换行、{len(texts)} 行标注）')
