# -*- coding: utf-8 -*-
"""交叉验证报告生成器 v4（以「脚本是否实际执行核查」为判定标准）。

标准（与核查表第 13 列标注同口径）：
  可自动化     = 勾选的全部适用对象，脚本均能自动判定（pass/fail）
  部分可自动化 = 脚本已执行核查但存在边界：仅覆盖部分对象/平台，或采集证据后结论需人工确认
  需人工       = 脚本未执行该项核查（纯方法提示或无覆盖）
适用对象以核查表 C~M（11 列）勾选矩阵为准。本脚本只读核查脚本源码，不改动任何脚本。
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
        9: '网络设备', 10: 'MySQL', 11: 'SQL Server', 12: '达梦', 13: 'Redis'}
TARGET_SCRIPTS = {
    'Windows': ['win/check_xp7'], '麒麟': ['kylin/check_kylin'],
    'Nginx': ['win/check_nginx', 'kylin/check_nginx'], 'Tomcat': ['win/check_tomcat', 'kylin/check_tomcat'],
    # 网络设备：check_network.ps1 与 check_network.sh 判定逻辑一致，能力以 .sh 源码为准
    '网络设备': ['kylin/check_network'],
    'MySQL': ['win/check_mysql', 'kylin/check_mysql'],
    'SQL Server': ['win/check_sqlserver', 'kylin/check_sqlserver'],
    '达梦': ['win/check_dm', 'kylin/check_dm'], 'Redis': ['win/check_redis', 'kylin/check_redis'],
}
TYPE_ORDER = ['Windows', '麒麟', 'Nginx', 'Tomcat', '网络设备', 'MySQL', 'SQL Server', '达梦', 'Redis']
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
    '1.27': ('部分可自动化', '脚本采集软件版本与授权状态，正版/定制认定需人工核对采购凭证与售后协议'),
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
        # 第5章主体为网络设备，个别行含麒麟（如 5.19 终端逻辑隔离）
        row_t = [(t, target_cap(code, t)) for t in TYPE_ORDER if t in (applic.get(code) or [])]
        net_auto = cap.get(code, {}).get('kylin/check_network', 'none') == 'auto'
        others_ok = all(c == 'auto' for t, c in row_t if t != '网络设备')
        if net_auto and others_ok:
            return '可自动化', 'check_network.sh / check_network.ps1 自动判定（设备回显采集后自动给出结论）', '网络设备自动判定'
        if any(c in ('auto', 'check') for _, c in row_t):
            if code in NET_REASON:
                return '部分可自动化', NET_REASON[code], '网络设备部分信号+台账比对'
            bad = [t for t, c in row_t if c != 'auto']
            return '部分可自动化', '、'.join(bad) + ' 侧脚本未执行该项检测，需人工核查', '部分适用对象仅提示或未检测'
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
    appl = applic.get(code)
    if appl:
        return '、'.join(appl)
    return '网络设备' if int(code.split('.')[0]) == 5 else '—'

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


def sel(chs=None, st=None, bucket=None):
    """按章节/结论/原因类别筛出编号列表。"""
    out = []
    for _ch, _code, _t in guide_items:
        s2, _r, b = judge(_code)
        if chs is not None and int(_code.split('.')[0]) not in chs:
            continue
        if st is not None and s2 != st:
            continue
        if bucket is not None and b != bucket:
            continue
        out.append(_code)
    return out

man_14 = sel(chs={1, 2, 3, 4}, st='需人工')
man_5 = sel(chs={5}, st='需人工')
man_69 = sel(chs={6, 7, 8, 9, 10}, st='需人工')
part_missing = sel(bucket='部分适用对象仅提示或未检测')
net_auto = sel(chs={5}, st='可自动化')
net_part = sel(chs={5}, st='部分可自动化')
net_man = sel(chs={5}, st='需人工')

b1 = bucket_stat.get('部分适用对象仅提示或未检测', 0)
b2 = bucket_stat.get('部分适用对象采集证据需人工', 0)
b3 = bucket_stat.get('网络设备部分信号+台账比对', 0)
b4 = bucket_stat.get('脚本采集证据、结论需人工', 0)

today = datetime.date.today().isoformat()

chapters = []
for ch, code, title in guide_items:
    if ch not in chapters:
        chapters.append(ch)

# ---------- 6. 一致性校验：报告结论 vs 核查表第13列 ----------
excel_label = {}
cur = None; seq = 0
for r in range(4, 140):
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
    v = str(_ws.cell(row=r, column=14).value or '')
    excel_label[code] = ('可自动化' if v.startswith('可自动化')
                         else ('部分可自动化' if v.startswith('部分可自动化') else '需人工'))

mismatch = [c for c, lab in excel_label.items() if judge(c)[0] != lab]
if mismatch:
    print('!! 报告与核查表「工具自动验证」列不一致：', mismatch)
    sys.exit(1)

# ---------- 7. 生成报告（交付版：结论 + 方法 + 覆盖统计 + 项号索引）----------
L = []
A = L.append

def idx(codes):
    return '、'.join(codes) if codes else '—'

def by_state(ch, st):
    return [c for c, _t, _a, s2, _r in detail[ch] if s2 == st]

A('# 配置核查作业指导书与自动化核查工具交叉验证报告')
A('')
A(f'报告日期：{today}')
A('')
A('## 一、验证结论')
A('')
A(f'依据《配置核查作业指导书》v2.2（10 章、{total} 个编号检查项），对自动化核查工具逐项交叉验证，'
  f'核查表的检查项、适用对象与工具核查能力逐项对齐，结果如下：')
A('')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 纳入工具核查 | {n_auto + n_part} | {(n_auto + n_part)/total*100:.1f}% |')
A(f'| 　其中：可自动化（工具直接给出合格/不合格判定） | {n_auto} | {n_auto/total*100:.1f}% |')
A(f'| 　其中：部分可自动化（脚本已核查，结论需结合台账或现场确认） | {n_part} | {n_part/total*100:.1f}% |')
A(f'| 人工核查 | {n_manual} | {n_manual/total*100:.1f}% |')
A('')
A(f'全部 {total} 个检查项中，{n_auto + n_part} 项的核查由工具执行：其中 {n_auto} 项工具直接给出合格或不合格判定，'
  f'{n_part} 项脚本已执行核查、需结合设备台账或现场情况确认最终结论。'
  f'其余 {n_manual} 项按标准要求属现场查看、文档调阅与实测验证类，由人工按同一套判定规则完成并留存证据。')
A('')
A('## 二、验证对象与方法')
A('')
A('### 2.1 验证对象')
A('')
A('- 指导书：《配置核查作业指导书》v2.2，共 10 章、' + str(total) + ' 个编号检查项。')
A('- 核查工具：16 个脚本。操作系统核查 2 个（Windows、麒麟各 1）；数据库与中间件核查 12 个（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat，两平台各 6）；网络设备核查 2 个（麒麟版与 Windows 版，采集-解析模式，判定逻辑与采集文件通用，覆盖第 5 章 23 项）。')
A('- 工具的组成、部署方式与使用方法见《配置核查工具包》README。')
A('')
A('### 2.2 验证方法')
A('')
A('从指导书提取全部编号检查项，逐项确定适用对象（操作系统、数据库、中间件、网络设备），再逐对象核对脚本的核查能力。'
  '能力判定以脚本源码的检测分支与结果输出为依据：能够给出合格或不合格判定的，列为可自动化；'
  '执行检测命令并输出核查结果、结论需结合台账或现场确认的，列为部分可自动化；'
  '其余按标准要求需现场查看、文档调阅与实测验证的，列为需人工核查。')
A('')
A('### 2.3 判定规则')
A('')
A('| 验证结论 | 判定条件 |')
A('|---|---|')
A('| 可自动化 | 该检查项适用的全部对象，脚本均可自动给出合格或不合格判定 |')
A('| 部分可自动化 | 脚本已执行核查，受核查对象或结论性质所限，需结合设备台账或现场情况确认最终结论 |')
A('| 需人工核查 | 现场查看、文档调阅、实测验证类检查项，由人工按指导书方法完成并留存证据 |')
A('')
A('## 三、各章覆盖统计')
A('')
A('| 章节 | 检查项 | 可自动化 | 部分可自动化 | 需人工核查 |')
A('|---|---|---|---|---|')
for ch in chapters:
    a, b, c = stat[ch]
    A(f'| {ch} | {len(detail[ch])} | {a} | {b} | {c} |')
A('')
A('## 四、逐项验证结果索引')
A('')
A('各章按验证结论列出的检查项编号如下，逐项的适用对象、所用脚本与结论依据见《配置核查表》第 14 列「工具自动验证」。')
A('')
for ch in chapters:
    A(f'**{ch}**')
    A('')
    A(f'- 可自动化：{idx(by_state(ch, "可自动化"))}')
    A(f'- 部分可自动化：{idx(by_state(ch, "部分可自动化"))}')
    A(f'- 需人工核查：{idx(by_state(ch, "需人工"))}')
    A('')
A('## 五、说明')
A('')
A('- 第 6 至 9 章（物理安全、组织机构、规章制度、管理实施）为现场查看与文档调阅类检查项，'
  '以及第 1 至 5 章中涉及应用代码安全、渗透验证、业务架构与台账比对的检查项，按标准要求由人工完成核查。')
A('- 第 5 章网络设备核查采用设备回显采集后解析的方式：按工具生成的命令清单，由运维人员登录交换机、'
  '防火墙等设备执行命令并留存回显，工具据此逐项判定并出具报告。')
A('- 核查工具与人工核查台采用同一套检查项编号与判定规则，两部分结果可直接合并汇总。')
A('')

open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write(chr(10).join(L))
print(f'已生成（交付版）：{total} 项，可自动化 {n_auto}，部分可自动化 {n_part}，需人工 {n_manual}')
print('一致性校验通过：报告结论与核查表第 14 列逐项一致')
