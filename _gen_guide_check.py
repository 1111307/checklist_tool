# -*- coding: utf-8 -*-
"""交叉验证报告生成器 v4（以「脚本是否实际执行核查」为判定标准）。

标准（与核查表第 13 列标注同口径）：
  可自动化     = 勾选的全部适用对象，脚本均能自动判定（pass/fail）
  部分可自动化 = 脚本已执行核查但存在边界：仅覆盖部分对象/平台，或采集证据后结论需人工确认
  需人工       = 脚本未执行该项核查（纯方法提示或无覆盖）
适用对象以核查表 C~L 勾选矩阵为准。本脚本只读核查脚本源码，不改动任何脚本。
"""
import re, glob, os, sys, datetime
from docx import Document
from openpyxl import load_workbook
from collections import defaultdict

BSLASH = chr(92)

# ---------- 1. 指导书检查项 ----------
doc = Document('配置核查作业指导书_v2.2.docx')
guide_items = []
cur_ch = None
for p in doc.paragraphs:
    if p.style.name == 'Heading 1':
        cur_ch = p.text.strip()
    elif p.style.name == 'Heading 2':
        m = re.match(r'(\d+\.\d+)\s+(.*)', p.text.strip())
        if m:
            guide_items.append((cur_ch, m.group(1), m.group(2)))

# ---------- 2. 脚本能力提取（与 _annotate_xlsx.py 同逻辑） ----------
ST_SH = re.compile(r'add_result\s+"([^"]+)"\s+"[^"]*"\s+"[^"]*"\s+"([^"]+)"')
ST_VBS = re.compile(r'AddResult\s*\(?\s*"([^"]+)"\s*,\s*"[^"]*"\s*,\s*"[^"]*"\s*,\s*"([^"]+)"')
EVIDENCE = re.compile(
    r'\$\{?\w|\$\(|"\s*&|检测到|未检测到|采集到|已安装|已部署|已启用|未发现|已划分|'
    r'执行失败|无法读取|无法查询|无法确认|无法获取|无法确定|权限不足|已建立|'
    r'未找到|未运行|未在|（输出|输出:')

def read_detail(data, pos, is_vbs):
    i = pos
    if is_vbs:
        m = re.match(r'\s*,\s*', data[i:i+20])
        if not m:
            return '', False
        i += m.end()
        if i >= len(data) or data[i] != '"':
            return '', False
        i += 1
        out = []
        while i < len(data):
            ch = data[i]
            if ch == '"':
                if i + 1 < len(data) and data[i+1] == '"':
                    out.append('"'); i += 2; continue
                break
            out.append(ch); i += 1
        concat = bool(re.match(r'\s*&', data[i+1:i+8])) if i + 1 < len(data) else False
        return ''.join(out), concat
    m = re.match(r'\s*\\?\s*', data[i:i+40], re.S)
    if m:
        i += m.end()
    if i >= len(data) or data[i] != '"':
        return '', False
    i += 1
    out = []
    while i < len(data):
        ch = data[i]
        if ch == BSLASH:
            out.append(data[i:i+2]); i += 2; continue
        if ch == '"':
            break
        out.append(ch); i += 1
    return ''.join(out), False

def extract(path, is_vbs):
    data = open(path, encoding='gbk' if is_vbs else 'utf-8').read()
    if is_vbs:
        data = re.sub(r'_\s*\r?\n\s*', '', data)
    st = ST_VBS if is_vbs else ST_SH
    d = defaultdict(lambda: {'statuses': set(), 'manuals': []})
    for m in st.finditer(data):
        code, status = m.group(1), m.group(2)
        d[code]['statuses'].add(status)
        if status == 'manual':
            det, concat = read_detail(data, m.end(), is_vbs)
            d[code]['manuals'].append((det, concat))
    return d

def script_cap(info):
    kinds = set(info['statuses']) - {'na'}
    if 'pass' in kinds or 'fail' in kinds:
        return 'auto'
    if 'manual' in kinds:
        for det, concat in info['manuals']:
            if EVIDENCE.search(det) or concat:
                return 'check'
        return 'reminder'
    if 'na' in info['statuses']:
        return 'na'
    return 'none'

scripts = {}
for f in sorted(glob.glob('win/check_*.vbs')):
    scripts['win/' + os.path.basename(f).replace('.vbs', '')] = extract(f, True)
for f in sorted(glob.glob('kylin/check_*.sh')):
    scripts['kylin/' + os.path.basename(f).replace('.sh', '')] = extract(f, False)

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

cap = defaultdict(dict)
ORDER = {'auto': 4, 'check': 3, 'reminder': 2, 'na': 1, 'none': 0}
for sname, d in scripts.items():
    for code, info in d.items():
        k = script_cap(info)
        for c in expand(code):
            if ORDER[k] > ORDER[cap[c].get(sname, 'none')]:
                cap[c][sname] = k

# ---------- 3. 适用矩阵 ----------
XLSX = '配置核查表_v2.0.0_标注自动验证.xlsx'
_wb = load_workbook(XLSX)
_ws = _wb.active
COLS = {3: 'Windows', 4: 'Windows', 5: '麒麟', 6: '麒麟', 7: 'Nginx', 8: 'Tomcat',
        9: 'MySQL', 10: 'SQL Server', 11: '达梦', 12: 'Redis'}
TARGET_SCRIPTS = {
    'Windows': ['win/check_xp7'], '麒麟': ['kylin/check_kylin'],
    'Nginx': ['win/check_nginx', 'kylin/check_nginx'], 'Tomcat': ['win/check_tomcat', 'kylin/check_tomcat'],
    'MySQL': ['win/check_mysql', 'kylin/check_mysql'],
    'SQL Server': ['win/check_sqlserver', 'kylin/check_sqlserver'],
    '达梦': ['win/check_dm', 'kylin/check_dm'], 'Redis': ['win/check_redis', 'kylin/check_redis'],
}
chapters_hdr = {'系统安全': 1, '用户安全': 2, '数据安全': 3, '应用安全': 4, '网络安全': 5, '物理安全': 6}
applic = {}
cur = None; seq = 0
for r in range(4, _ws.max_row + 1):
    c1 = _ws.cell(row=r, column=1).value
    if c1 and str(c1).strip() in chapters_hdr:
        cur = chapters_hdr[str(c1).strip()]; seq = 0
    seq += 1
    if cur == 6:
        if seq <= 9: code = f'6.{seq}'
        elif seq <= 14: code = f'7.{seq-9}'
        elif seq <= 18: code = f'8.{seq-14}'
        elif seq <= 22: code = f'9.{seq-18}'
        else: code = '10.1'
    else:
        code = f'{cur}.{seq}'
    if r <= 138:
        appl = []
        for c, tname in COLS.items():
            if str(_ws.cell(row=r, column=c).value or '').strip() == '√' and tname not in appl:
                appl.append(tname)
        applic[code] = appl

def target_cap(code, t):
    best = 'none'
    for s in TARGET_SCRIPTS[t]:
        k = cap.get(code, {}).get(s, 'none')
        if k == 'auto':
            return 'auto'
        if k == 'check' and best != 'auto':
            best = 'check'
        if k == 'reminder' and best in ('none',):
            best = 'reminder'
    return best

# ---------- 4. 判定 ----------
SPECIAL = {
    '10.1': ('可自动化', '脚本自动核查 TLS/SSH 协议版本（协议层条目，勾选矩阵未单列对象）'),
    '1.27': ('部分可自动化', '脚本采集版本与授权状态，正版/定制认定需人工核对采购凭证（核查表未设行）'),
}
NET_REASON = {
    '5.9': 'check_network.sh 提取设备型号，异构部署结论需比对台账',
    '5.12': 'check_network.sh 采集 IKE SA 识别网络层加密，加密层次组合需按方案人工核对',
    '5.14': 'check_network.sh 核对防火墙会话表比对跨域通道，隔离方式与等级对应需人工核查',
    '5.17': 'check_network.sh 采集 IKE SA 识别网络层加密，链路层/信源层加密需按方案人工核对',
    '5.22': 'check_network.sh 提取存储网/业务网 VLAN 划分信号，结论需比对台账',
}

def judge(code):
    """返回 (结论, 原因, 原因类别)"""
    ch = int(code.split('.')[0])
    if code in SPECIAL:
        st, reason = SPECIAL[code]
        return st, reason, ('脚本采集证据、结论需人工' if st == '部分可自动化' else '特例')
    if ch == 5:
        k = cap.get(code, {}).get('kylin/check_network', 'none')
        if k == 'auto':
            return '可自动化', 'check_network.sh 自动判定（设备回显采集后自动给出结论）', '网络设备自动判定'
        if k == 'check':
            return '部分可自动化', NET_REASON.get(code, '脚本采集部分信号，结论需人工'), '网络设备部分信号+台账比对'
        return '需人工', '审批/台账/平台类核查，脚本未执行检测', '第5章台账/平台类'
    if ch in (6, 7, 8, 9):
        return '需人工', '现场/文档/台账类核查', '第6-9章现场/文档类'
    appl = applic.get(code)
    if not appl:
        covered = any(cap.get(code, {}).get(s, 'none') in ('auto', 'check')
                      for s in ('win/check_xp7', 'kylin/check_kylin'))
        if covered:
            return '需人工', '未勾选核查对象（如载体销毁类条目），脚本未执行检测', '第1-4章脚本未检测'
        return '需人工', '脚本未执行该项检测', '第1-4章脚本未检测'
    auto_t = [t for t in appl if target_cap(code, t) == 'auto']
    check_t = [t for t in appl if target_cap(code, t) == 'check']
    rem_t = [t for t in appl if target_cap(code, t) not in ('auto', 'check')]
    if not check_t and not rem_t:
        return '可自动化', '适用对象均自动判定', '适用对象均自动判定'
    if not auto_t and not check_t:
        return '需人工', '脚本未执行该项检测，仅输出指导书方法提示', '第1-4章脚本未检测'
    # 脚本已执行核查但存在边界 → 部分可自动化
    parts = []
    if auto_t:
        parts.append('、'.join(auto_t) + ' 自动判定')
    if check_t:
        parts.append('、'.join(check_t) + ' 脚本采集证据、结论需人工')
    if rem_t:
        parts.append('、'.join(rem_t) + ' 未检测需人工')
    reason = '；'.join(parts)
    if rem_t:
        bucket = '部分适用对象仅提示或未检测'
    else:
        bucket = '部分适用对象采集证据需人工'
    return '部分可自动化', reason, bucket
    return '需人工', '脚本未执行该项检测，仅输出指导书方法提示', '第1-4章脚本未检测'

def applic_disp(code):
    ch = int(code.split('.')[0])
    if ch == 5:
        return '网络设备'
    appl = applic.get(code)
    if not appl:
        return '—'
    return '、'.join(appl)

# ---------- 5. 统计 ----------
stat = defaultdict(lambda: [0, 0, 0])
bucket_stat = defaultdict(int)
detail = defaultdict(list)
for ch, code, title in guide_items:
    st, reason, bucket = judge(code)
    if st == '可自动化':
        stat[ch][0] += 1
    elif st == '部分可自动化':
        stat[ch][1] += 1; bucket_stat[bucket] += 1
    else:
        stat[ch][2] += 1
    detail[ch].append((code, title, applic_disp(code), st, reason))

total = len(guide_items)
n_auto = sum(v[0] for v in stat.values())
n_part = sum(v[1] for v in stat.values())
n_manual = sum(v[2] for v in stat.values())
rate = n_auto / total * 100
today = datetime.date.today().isoformat()

chapters = []
for ch, code, title in guide_items:
    if ch not in chapters:
        chapters.append(ch)

# ---------- 6. 一致性校验：报告结论 vs 核查表第13列 ----------
excel_label = {}
cur = None; seq = 0
for r in range(4, 139):
    c1 = _ws.cell(row=r, column=1).value
    if c1 and str(c1).strip() in chapters_hdr:
        cur = chapters_hdr[str(c1).strip()]; seq = 0
    seq += 1
    if cur == 6:
        if seq <= 9: code = f'6.{seq}'
        elif seq <= 14: code = f'7.{seq-9}'
        elif seq <= 18: code = f'8.{seq-14}'
        elif seq <= 22: code = f'9.{seq-18}'
        else: code = '10.1'
    else:
        code = f'{cur}.{seq}'
    v = str(_ws.cell(row=r, column=13).value or '')
    excel_label[code] = ('可自动化' if v.startswith('可自动化')
                         else ('部分可自动化' if v.startswith('部分可自动化') else '需人工'))

mismatch = [c for c, lab in excel_label.items() if judge(c)[0] != lab]
if mismatch:
    print('!! 报告与核查表第13列不一致：', mismatch)
    sys.exit(1)

# ---------- 7. 生成报告 ----------
L = []
A = L.append

A('# 配置核查指导书与自动化核查工具交叉验证报告')
A('')
A(f'报告日期：{today}')
A('')
A('## 一、验证结论')
A('')
A(f'对《配置核查作业指导书》v2.2 与自动化核查工具（Windows VBScript 版、麒麟 Bash 版、网络设备核查 check_network.sh）逐项比对，共核对 {total} 个编号检查项。以「脚本是否实际执行核查」为判定标准，结合核查表勾选矩阵（√ 适用 / — 不适用）逐对象核对脚本能力：可自动化 {n_auto} 项，部分可自动化 {n_part} 项，需人工 {n_manual} 项，自动化覆盖率 {rate:.1f}%。')
A('')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 可自动化 | {n_auto} | {n_auto/total*100:.1f}% |')
A(f'| 部分可自动化 | {n_part} | {n_part/total*100:.1f}% |')
A(f'| 需人工 | {n_manual} | {n_manual/total*100:.1f}% |')
A('')
b1 = bucket_stat.get('部分适用对象仅提示或未检测', 0)
b2 = bucket_stat.get('部分适用对象采集证据需人工', 0)
b3 = bucket_stat.get('网络设备部分信号+台账比对', 0)
b4 = bucket_stat.get('脚本采集证据、结论需人工', 0)
A(f'部分可自动化 {n_part} 项均为脚本已执行核查但存在边界：部分适用对象仅提示或未检测 {b1} 项、部分对象采集证据后结论需人工 {b2} 项、网络设备部分信号需台账比对 {b3} 项、版本采集比对 {b4} 项。需人工 {n_manual} 项为脚本未执行该项核查的条目（第 1-4 章 17 项、第 5 章台账/平台类 6 项、第 6-9 章现场文档类 22 项）。')
A('')
A('## 二、验证对象与方法')
A('')
A('### 2.1 验证对象')
A('')
A('- 指导书：《配置核查作业指导书》v2.2，共 10 章、' + str(total) + ' 个编号检查项（交付包内的 v2.0.1 为同内容版本，另含第 5 章 13 张终端操作截图）。')
A('- 核查工具：16 个脚本。操作系统核查 2 个（Windows、麒麟各 1），数据库与中间件核查 12 个（MySQL、Redis、达梦 DM、SQL Server、Nginx、Tomcat，两平台各 6），网络设备核查 2 个（麒麟 check_network.sh 与 Windows check_network.ps1，采集-解析模式、判定逻辑一致、采集文件通用，覆盖第 5 章 23 项）。')
A('')
A('### 2.2 验证方法')
A('')
A('从指导书 docx 提取全部编号检查项，从各脚本源码提取检查结果记录点。每项的适用对象以核查表勾选矩阵为准（√ 适用 / — 不适用）：Windows/WinXP 对应 check_xp7.vbs、中标/银河麒麟对应 check_kylin.sh、各数据库与中间件对应双平台组件脚本、网络设备对应 check_network.sh。「脚本是否实际执行核查」以脚本源码的检测分支与结果详情为依据：有自动判定（合格/不合格），或执行检测命令并输出采集结果（结论留人工确认）的，均属已核查；仅输出方法提示、未执行任何检测的，属未核查。')
A('')
A('### 2.3 判定规则')
A('')
A('| 验证结论 | 判定条件 |')
A('|---|---|')
A('| 可自动化 | 该项在核查表中勾选的全部适用对象，核查脚本均能自动给出合格或不合格判定 |')
A('| 部分可自动化 | 脚本已执行核查但存在边界：仅覆盖部分适用对象或一侧平台，或采集证据后结论需人工确认 |')
A('| 需人工 | 脚本未执行该项核查（仅输出指导书方法提示或无覆盖），需人工完成 |')
A('')
A('## 三、各章覆盖统计')
A('')
A('| 章节 | 检查项 | 可自动化 | 部分可自动化 | 需人工 |')
A('|---|---|---|---|---|')
for ch in chapters:
    a, b, c = stat[ch]
    A(f'| {ch} | {len(detail[ch])} | {a} | {b} | {c} |')
A('')
A('## 四、逐项对照明细')
A('')
for ch in chapters:
    A(f'### {ch}')
    A('')
    A('| 编号 | 检查项 | 适用对象 | 验证结论 | 原因 |')
    A('|---|---|---|---|---|')
    for code, title, appl, st, reason in detail[ch]:
        A(f'| {code} | {title} | {appl} | {st} | {reason} |')
    A('')

A('## 五、差异与问题分析')
A('')
A('### 5.1 网络安全（第 5 章）')
A('')
A('第 5 章 23 项分三类：设备命令回显类 12 项（5.3 至 5.8、5.10、5.11、5.16、5.18、5.19、5.23，其中 5.23 为端口闲置率自动判定）由 check_network.sh 自动判定；5 项脚本采集部分信号、结论需人工（5.9 设备型号、5.12/5.17 IKE SA 网络层加密、5.14 防火墙会话表、5.22 VLAN 划分）；6 项为审批记录、管理平台与方案文档类（5.1、5.2、5.13、5.15、5.20、5.21），脚本未执行检测。该划分与核查表第 13 列「工具自动验证」标注一一对应。')
A('')
A('### 5.2 部分可自动化项的原因分布')
A('')
A(f'共 {n_part} 项为部分可自动化，按原因分布如下：')
A('')
A('| 原因类别 | 项数 |')
A('|---|---|')
for b, c in sorted(bucket_stat.items(), key=lambda x: -x[1]):
    A(f'| {b} | {c} |')
A('')
A('### 5.3 需人工项的构成')
A('')
A(f'共 {n_manual} 项需人工：第 6 至 9 章（物理、组织、制度、管理）22 项，要到机房看设备台账、翻组织文件、核对制度版本和演练记录；第 5 章台账/平台类 6 项，是跨网审批单、方案文档与管理平台界面核查；第 1-4 章 17 项，其中应用层代码能力 8 项（1.16 输入参数校验、4.9、4.10、4.15、4.20、4.21、4.26 至 4.29）需渗透测试、代码审计或 WAF 工具，物理过程与文档台账比对 9 项（1.13、1.22、3.4、3.12、3.14、4.5、4.13 等）需人工核对业务架构、销毁档案与边界设备。')
A('')
A(f'部分可自动化 {n_part} 项里，{b1} 项存在适用对象仅提示或未检测（麒麟侧 1.25、2.8、2.15、4.12、4.23、4.24，Windows 侧 4.31，组件侧 1.20、1.21、4.1，跨对象 1.24），后续补齐对应判定分支即可转可自动化；其余为采集证据后结论本需人工比对（达梦版本比对、备份有效性验证、网络设备台账比对等），属合理边界。本报告仅核对脚本现状，未改动任何脚本代码。')
A('')
A('## 六、整改建议')
A('')
A('下一步几件事。第 5 章上真机核查时，用 check_network.sh init 生成采集清单，运维陪同采集回显后跑 check 出报告；遇到解析不出的回显格式，把回显片段补进特征匹配。需人工的 45 项用人工核查台（manual_check.html）逐项记录、粘贴取证截图并导出报告。如需提升自动化率，可补齐一侧平台与组件对象的判定分支（约 10 项可转可自动化），补齐前以本报告口径为准。')
A('')

open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write('\n'.join(L))
print(f'已生成：{total} 项，可自动化 {n_auto}，部分可自动化 {n_part}，需人工 {n_manual}，覆盖率 {rate:.1f}%')
print('原因分布：', dict(bucket_stat))
print('一致性：报告与核查表第13列逐项一致 ✓')
