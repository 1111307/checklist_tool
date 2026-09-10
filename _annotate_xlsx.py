# -*- coding: utf-8 -*-
"""核查表第13列「工具自动验证」详情标注 v2（按适用对象矩阵）。

判定基准：C~L 列勾选矩阵（√=适用、—=不适用）。逐行检查每个适用对象：
  - Windows/麒麟 → check_xp7.vbs / check_kylin.sh 是否有 pass/fail 自动分支
  - 数据库/中间件 → 双平台组件脚本任一有 pass/fail 即该对象可自动
适用对象全部可自动 → 可自动化；部分可自动 → 部分可自动化（写明哪些能、哪些不能）。
分类与交叉验证报告同口径（报告同日按矩阵逻辑重生成）。
用法：python _annotate_xlsx.py [--apply]
"""
import re, glob, os, sys, math
from collections import defaultdict, Counter
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
    out.add(code); return out

cov = defaultdict(lambda: defaultdict(set))
for sname, d in scripts.items():
    for code, sts in d.items():
        for c in expand(code):
            cov[c][sname] |= sts

# ---------- 2. 适用矩阵 ----------
wb = load_workbook(XLSX)
ws = wb.active
COLS = {3: 'Windows', 4: 'Windows', 5: '麒麟', 6: '麒麟', 7: 'Nginx', 8: 'Tomcat',
        9: 'MySQL', 10: 'SQL Server', 11: '达梦', 12: 'Redis'}
TARGET_SCRIPTS = {
    'Windows': ['win/check_xp7'], '麒麟': ['kylin/check_kylin'],
    'Nginx': ['win/check_nginx', 'kylin/check_nginx'], 'Tomcat': ['win/check_tomcat', 'kylin/check_tomcat'],
    'MySQL': ['win/check_mysql', 'kylin/check_mysql'],
    'SQL Server': ['win/check_sqlserver', 'kylin/check_sqlserver'],
    '达梦': ['win/check_dm', 'kylin/check_dm'], 'Redis': ['win/check_redis', 'kylin/check_redis'],
}
PROD_ORDER = ['MySQL', 'Redis', '达梦', 'SQL Server', 'Nginx', 'Tomcat']

chapters = {'系统安全': 1, '用户安全': 2, '数据安全': 3, '应用安全': 4, '网络安全': 5, '物理安全': 6}
rowmap, applic_map = {}, {}
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
    if r <= 138:
        appl = []
        for c, tname in COLS.items():
            if str(ws.cell(row=r, column=c).value or '').strip() == '√' and tname not in appl:
                appl.append(tname)
        applic_map[code] = appl

def target_capability(code, t):
    """返回 'auto' / 'manual' / 'none'（按该对象对应脚本的最佳能力）"""
    best = 'none'
    for s in TARGET_SCRIPTS[t]:
        sts = cov.get(code, {}).get(s, set())
        if 'pass' in sts or 'fail' in sts:
            return 'auto'
        if 'manual' in sts:
            best = 'manual'
    return best

# ---------- 3. 手工文案（适用对象/脚本详情已逐项核对） ----------
HAND = {
    # 全人工但有脚本方法（适用对象均无自动分支）
    '1.13': '部分可自动化（脚本输出数据库存储配置，分类独立存储与业务符合性需人工核对）',
    '1.16': '部分可自动化（数据库脚本输出输入检查要点，参数化查询等应用层校验需 sqlmap/AWVS 或人工核查）',
    '1.22': '部分可自动化（脚本可查审计开关与参数，独立安全监护与审计措施需人工确认）',
    '3.4':  '部分可自动化（载体销毁为物理过程，脚本不适用；需人工核查销毁档案与影像记录）',
    '3.12': '部分可自动化（脚本可查采集/同步服务进程，采集范围与业务需求比对需人工）',
    '3.14': '部分可自动化（OS 层权限与审计可查，大数据平台统一管控需平台界面人工核查）',
    '4.1':  '部分可自动化（OS 定时备份任务可查，应用自身备份功能需登录应用核查）',
    '4.5':  '部分可自动化（脚本可查服务与限流配置线索，防DDoS多层协同能力需人工/边界设备核查）',
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
    '4.32': '部分可自动化（脚本检测备份目录与备份文件，恢复能力需人工验证）',
    # 全自动但写明具体检测内容（信息量优于模板）
    '2.3':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查自动登录与空口令账户限制）',
    '2.7':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查系统盘文件保护状态）',
    '3.6':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 DLP/终端管控客户端部署）',
    '3.1':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查磁盘加密 BitLocker/LUKS）',
    '10.1': '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 TLS/SSH 协议版本；协议层条目，勾选矩阵未单列对象）',
    # 第5章 部分（脚本提取信号 + 台账）
    '5.9':  '部分可自动化（用 check_network.sh 提取设备型号，异构部署结论需比对台账）',
    '5.22': '部分可自动化（用 check_network.sh 提取存储网/业务网划分信号，结论需比对台账）',
    '5.23': '部分可自动化（用 check_network.sh 计算端口闲置率，配备必要性结论需人工评估）',
}
NET_AUTO = {'5.3', '5.4', '5.5', '5.6', '5.7', '5.8', '5.10', '5.11', '5.16', '5.18', '5.19'}

def disp(t, role):
    if t in ('Windows', '麒麟') and role == 'manual':
        return f'{t} OS 层'
    return t

def compose(code):
    ch = int(code.split('.')[0])
    if code in HAND:
        return HAND[code]
    if ch == 5:
        if code in NET_AUTO:
            return '可自动化（用 check_network.sh 采集设备回显后自动判定）'
        return '需人工（台账/平台核查）'
    if ch in (6, 7, 8, 9):
        return '需人工（现场/文档/台账类核查）'
    # 第1-4章、10.1：按适用矩阵
    appl = applic_map.get(code, [])
    auto_os = [t for t in ('Windows', '麒麟') if t in appl and target_capability(code, t) == 'auto']
    manual_os = [t for t in ('Windows', '麒麟') if t in appl and target_capability(code, t) != 'auto']
    auto_prods = [t for t in PROD_ORDER if t in appl and target_capability(code, t) == 'auto']
    manual_prods = [t for t in PROD_ORDER if t in appl and target_capability(code, t) != 'auto']
    if not appl:
        return None  # 不应出现（10.1 已在 HAND）
    if not manual_os and not manual_prods:
        # 全部适用对象可自动
        os_part = '用 check_xp7.vbs、check_kylin.sh' if len(auto_os) == 2 else (
            '用 check_xp7.vbs' if auto_os == ['Windows'] else ('用 check_kylin.sh' if auto_os == ['麒麟'] else ''))
        prod_part = '及 ' + '、'.join(auto_prods) + ' 组件脚本' if auto_prods else ''
        if os_part:
            return f'可自动化（{os_part}{(" " + prod_part) if prod_part else ""} 自动核查）'
        return f'可自动化（用{" " + "、".join(auto_prods) + " 组件脚本自动核查" if auto_prods else ""}）'
    # 混合：写明可自动部分 + 需人工对象
    bits = []
    if auto_os:
        s = '用 check_xp7.vbs、check_kylin.sh' if len(auto_os) == 2 else (
            '用 check_xp7.vbs' if auto_os == ['Windows'] else '用 check_kylin.sh')
        bits.append(s)
    if auto_prods:
        bits.append('及 ' + '、'.join(auto_prods) + ' 组件脚本' if bits else '用 ' + '、'.join(auto_prods) + ' 组件脚本')
    manual_list = [disp(t, 'manual') for t in manual_os] + manual_prods
    if bits:
        return f'部分可自动化（{" ".join(bits)} 自动核查；{"、".join(manual_list)} 需人工）'
    return None  # 全人工行应走 HAND

# ---------- 4. 生成 + 校验 ----------
texts, errors = {}, []
dist = Counter()
for r in range(4, 139):
    code = rowmap[r]
    txt = compose(code)
    if txt is None:
        errors.append(f'R{r} {code}: 无文案（全人工行缺 HAND 词条）')
        continue
    # 一致性：可自动化 ⇔ 全部适用对象有自动分支（5.x 网络与 10.1 除外）
    if txt.startswith('可自动化') and code not in HAND and int(code.split('.')[0]) != 5:
        appl = applic_map.get(code, [])
        if any(target_capability(code, t) != 'auto' for t in appl):
            errors.append(f'R{r} {code}: 标可自动化但存在需人工适用对象')
    if txt.startswith('部分可自动化') and code not in HAND:
        appl = applic_map.get(code, [])
        if all(target_capability(code, t) == 'auto' for t in appl) and appl:
            errors.append(f'R{r} {code}: 标部分但全部适用对象可自动')
    texts[r] = txt
    dist['可自动化' if txt.startswith('可自动化') else ('部分可自动化' if txt.startswith('部分') else '需人工')] += 1

if errors:
    print('!! 校验失败：')
    for e in errors:
        print('  ', e)
    sys.exit(1)

EXPECT = {'可自动化': 60, '部分可自动化': 44, '需人工': 31}
print(f'=== 干跑通过：{len(texts)} 条，分布 {dict(dist)}，预期 {EXPECT} ===')
if dict(dist) != EXPECT:
    print('!! 分布与预期不符，中止（需同步更新预期值与报告口径）')
    sys.exit(1)
if not APPLY:
    print('（干跑模式，未写盘。加 --apply 应用）')
    sys.exit(0)

# ---------- 5. 写入 ----------
for r, txt in texts.items():
    ws.cell(row=r, column=13).value = txt
ws.column_dimensions['M'].width = 55
for r, txt in texts.items():
    c = ws.cell(row=r, column=13)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    lines = math.ceil(len(txt) / 26)
    need = {1: 17, 2: 34, 3: 50}.get(lines, 50)
    h = ws.row_dimensions[r].height
    if h is not None and h < need:
        ws.row_dimensions[r].height = need

wb.save(XLSX)
print(f'已写入 {XLSX}（按适用对象矩阵，{len(texts)} 行标注）')
