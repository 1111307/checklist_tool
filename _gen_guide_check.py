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
item_name = {}
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
    # 检查项名称取核查表 B 列正式条目文本（去掉行首 "1、" 序号，与编号列呼应）
    nm = str(_ws.cell(row=r, column=2).value or '').strip()
    nm = re.sub(r'^[0-9]+[、.]', '', nm).strip()
    nm = re.sub(r'\s+', ' ', nm).strip().replace('|', '｜')   # 单元格内换行会切断 md 表格
    item_name[code] = nm

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
A(f'本次验证以《配置核查作业指导书》v2.2（10 章、{total} 个编号检查项）为基准，'
  f'逐条比对自动化核查工具的核查能力，结果如下：')
A('')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 纳入工具核查 | {n_auto + n_part} | {(n_auto + n_part)/total*100:.1f}% |')
A(f'| 　其中：可自动化（工具直接给出合格/不合格判定） | {n_auto} | {n_auto/total*100:.1f}% |')
A(f'| 　其中：部分可自动化（脚本已核查，结论需结合台账或现场确认） | {n_part} | {n_part/total*100:.1f}% |')
A(f'| 人工核查 | {n_manual} | {n_manual/total*100:.1f}% |')
A('')
A(f'全部 {total} 个检查项中，{n_auto + n_part} 项由工具执行核查：{n_auto} 项工具直接给出合格或不合格判定，'
  f'{n_part} 项脚本已执行核查、结论需结合设备台账或现场情况确认。'
  f'其余 {n_manual} 项属现场查看、文档调阅或实测验证类，由人工按同一套判定规则完成并留存证据。')
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
A('从指导书提取全部编号检查项，先确定每项的适用对象（操作系统、数据库、中间件、网络设备），再按对象核对脚本的核查能力。'
  '判定依据是脚本源码里的检测分支和结果输出：能给出合格或不合格判定的列为可自动化；'
  '执行了检测命令、结论还需结合台账或现场确认的列为部分可自动化；'
  '其余属现场查看、文档调阅或实测验证类的列为需人工核查。')
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
A('## 四、逐项验证结果')
A('')
A('各章检查项及验证结论列示如下；每项的核查对象、所用脚本与结论依据见《配置核查表》第 14 列「工具自动验证」。')
A('')
for ch in chapters:
    A(f'### {ch}')
    A('')
    A('| 编号 | 检查项 | 验证结论 |')
    A('|---|---|---|')
    for code, title, appl, st, reason in detail[ch]:
        A(f'| {code} | {item_name.get(code) or title} | {st} |')
    A('')
A('## 五、说明')
A('')
A('| 事项 | 说明 |')
A('|---|---|')
A('| 人工核查范围 | 第 6 至 9 章（物理安全、组织机构、规章制度、管理实施）为现场查看与文档调阅类；第 1 至 5 章中涉及应用代码安全、渗透验证、业务架构与台账比对的检查项，按标准要求由人工完成核查。 |')
A('| 网络设备核查方式 | 第 5 章按工具生成的命令清单，由运维人员登录交换机、防火墙等设备执行并留存回显，工具据此逐条判定并出具报告。 |')
A('| 结果合并 | 核查工具与人工核查台采用同一套检查项编号与判定规则，两部分结果可直接合并汇总。 |')
A('')
A('## 六、验证证据')
A('')
A('以下截图取自核查工具实际运行的输出，对应第一节至第三节的验证结论。')
A('')
# 每张图的说明（环境、方法、关键数据），排在图片下方，便于甲方看图即懂
FIG_NOTES = {
    'term_win_os': '环境为本机 Windows 10 专业版。执行 cscript check_xp7.vbs 后，脚本逐项输出 139 项判定（[OK] 合规 / [!!] 不合规 / [??] 需人工核查 / [--] 不适用），并生成 HTML、XLS、XLSX 三种格式报告。',
    'term_win_net': '按「采集—解析」两阶段执行：init 生成设备命令清单，运维人员登录设备留存回显后，check 解析回显给出第 5 章 23 项结论；图中可见两阶段输出与报告落盘。',
    'term_kylin_os': '环境为银河麒麟 V10 靶机。执行 bash check_kylin.sh，输出 139 项判定汇总与三种格式报告路径。',

    'run_win_sqlserver': '本机连接 SQL Server 实例（1433 端口）实跑，脚本识别到版本 16.0.4265.3，18 项逐项判定后生成报告。',
    'run_win_mysql': '本机未部署 MySQL 服务，脚本连接失败后将该组件 19 项全部转入人工核查，并照常生成报告文件。',
    'run_win_redis': '本机未部署 Redis 服务，16 项转入人工核查。',
    'run_win_dm': '本机未安装达梦 disql 客户端，脚本无法执行检测，19 项标记为不适用。',
    'run_win_nginx': '本机存在 Nginx 服务，脚本完成 8 项核查并生成报告。',
    'run_win_tomcat': '本机存在 Tomcat 服务，脚本完成 8 项核查并生成报告。',

    'run_kylin_mysql': '银河麒麟靶机容器内连接 MariaDB 10.3.39 实跑，19 项判定后生成三种格式报告。',
    'run_kylin_redis': '银河麒麟靶机容器内 Redis 实跑，16 项判定后生成报告。',
    'run_kylin_nginx': '银河麒麟靶机容器内 Nginx 实跑，8 项判定后生成报告。',
    'run_kylin_tomcat': '银河麒麟靶机容器内 Tomcat 实跑，8 项判定后生成报告。',
    'run_kylin_dm': '靶机未安装达梦 disql 客户端，19 项标记为不适用。',
    'run_kylin_sqlserver': '靶机无 SQL Server 实例，18 项标记为不适用。',

    'fig_win_os': 'Windows 版操作系统核查报告，含四态统计卡与 139 项明细表，每项列出判定、检测详情与对应指导书条款。',
    'fig_win_sqlserver': 'SQL Server 组件核查报告，统计卡显示 18 项的判定分布，明细列出每项的检测详情与建议。',
    'fig_win_net': '网络设备核查报告，按设备回显逐项解析第 5 章 23 项，表中标注判定依据。',
    'fig_kylin_os': '麒麟版操作系统核查报告，系统信息区显示实际核查的发行版与内核。',
    'fig_kylin_mysql': 'MySQL 组件核查报告，报告头显示连接串、数据库版本与核查时间。',
    'fig_kylin_nginx': 'Nginx 中间件核查报告，8 项判定与对应指导书条款。',
    'fig_kylin_net': '网络设备核查报告，按设备回显解析第 5 章 23 项。',

    'fig_manual': '人工核查台的初始状态：136 项检查项按章分组，逐项提供合规/不合规/需人工核查/不适用四个选项与详情输入框，支持点击卡片后 Ctrl+V 粘贴截图取证，填写内容自动保存在本机浏览器。',
    'fig_manual_filled': '全部 136 项填写完成的状态（示例数据）：顶部统计显示未核查 0 项，各章标题右侧显示已核项数，每项均填有结论与核查详情。示例结论为 合规 121、不合规 8、需人工核查 6、不适用 1；实际核查时由核查人按现场情况逐项填写。',
    'fig_manual_export': '点击「导出报告 (Excel)」后由核查台生成的 Excel 报告，可在 Excel 中直接打开：表头为现场信息与四态统计，表体为 136 行 × 8 列（章节、编号、类别、核查项、结果、详情、建议、参考指导书），与自动脚本出具的报告同格式，可直接合并汇总。图中数据为上一张的示例填写结果。',
    'fig6_matrix': '核查表第 14 列逐项标注每个检查项用哪些脚本、哪些对象可自动、哪些需人工，与报告结论同源。',

    'probe_win_1_4': '本机 Windows 实际环境。上半为脚本输出中的该项判定行，下半为脚本源码里该项的原始检测命令 netsh advfirewall show allprofiles 的实际输出，显示防火墙已启用。',
    'probe_win_1_5': '本机 Windows 实际环境。下半为检测命令 reg query …NetBT\\Parameters\\Interfaces /s 的实际输出，接口 NetbiosOptions 为 0x0 即未禁用。',
    'probe_win_2_11': '本机 Windows 实际环境。下半为检测命令 net accounts 的实际输出，密码长度最小值 0、锁定阈值「从不」，均不满足要求。',
    'probe_win_1_2': '本机 Windows 实际环境。下半为脚本所用的 WMI 查询 AntiVirusProduct 的实际输出，列出奇安信天擎与 Windows Defender 两项。',
    'probe_win_2_10': '本机 Windows 实际环境。下半为检测命令 reg query …USBSTOR /v Start 的实际输出，Start=0x3 表示 USB 存储未禁用（要求为 4）。',
    'probe_win_2_12': '本机 Windows 实际环境。下半为检测命令 sc query 对 WLANSVC 与 bthserv 的实际输出，两项服务均处于 RUNNING。',
    'probe_win_1_3': '本机 Windows 实际环境。下半为 net start 列出的运行中服务，未见 telnet/tftp/ftp/snmptrap 类已知高危服务，故该脚本仅能判到「需人工核查」。',
    'probe_kylin_1_1': '麒麟靶机实际环境。下半为检测命令 yum check-update 的实际输出，逐条列出可更新软件包，累计 46 个。',
    'probe_kylin_1_5': '麒麟靶机实际环境。下半为检测命令 sestatus 的实际输出，SELinux 状态为 disabled。',
    'probe_kylin_1_6': '麒麟靶机实际环境。下半为检测命令 grep -i ciphers /etc/ssh/sshd_config 的实际输出，已显式配置 Ciphers 且不含已知弱算法。',
    'probe_mysql_1_7': '麒麟靶机 MySQL 实际环境。下半为查询 mysql.user 的实际输出，账户口令字段为空。',
    'probe_mysql_1_8': '麒麟靶机 MySQL 实际环境。下半为查询 information_schema.ROUTINES 与 mysql.func 的实际输出，存储过程/函数总数 2、无自定义 UDF。',
    'probe_mysql_1_9': '麒麟靶机 MySQL 实际环境。下半为查询危险全局权限账户的实际输出，checker@localhost 与 checker@% 两个非管理员账户持有 FILE/SUPER/GRANT 等权限。',
    'probe_mysql_1_10': '麒麟靶机 MySQL 实际环境。下半为 ss -lntp 过滤 3306 端口的实际输出，仅监听 127.0.0.1，未对外暴露。',
    'probe_mysql_1_12': '麒麟靶机 MySQL 实际环境。下半为查询审计相关系统变量的实际输出，general_log 为 ON。',
    'probe_redis_1_7': '麒麟靶机 Redis 实际环境。下半为检测命令 redis-cli CONFIG GET requirepass 的实际输出，口令值为空，说明未设置认证口令。',
    'probe_redis_1_10': '麒麟靶机 Redis 实际环境。下半为检测命令 redis-cli CONFIG GET bind 的实际输出，绑定地址为 127.0.0.1，访问受限。',
}

FIG_GROUPS = [
    ('6.1 操作系统与网络设备核查运行过程', None, [
        ('term_win_os', '图 1　Windows 版运行操作系统核查脚本（cscript check_xp7.vbs，139 项逐项输出判定并生成报告）'),
        ('term_win_net', '图 2　Windows 版网络设备核查（init 生成采集清单 → 采集 → check 解析出报告，第 5 章 23 项）'),
        ('term_kylin_os', '图 3　麒麟版运行操作系统核查脚本（bash check_kylin.sh，139 项）'),
    ]),
    ('6.2 Windows 平台各组件核查脚本运行过程', None, [
        ('run_win_sqlserver', '图 4　SQL Server 核查脚本（连接 1433 端口实跑，18 项：合规 4、不合规 4、需人工核查 10）'),
        ('run_win_mysql', '图 5　MySQL 核查脚本（本机未部署 MySQL，19 项全部转入人工核查）'),
        ('run_win_redis', '图 6　Redis 核查脚本（本机未部署 Redis，16 项全部转入人工核查）'),
        ('run_win_dm', '图 7　达梦核查脚本（本机无 disql 客户端，19 项标记为不适用）'),
        ('run_win_nginx', '图 8　Nginx 核查脚本（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('run_win_tomcat', '图 9　Tomcat 核查脚本（8 项：合规 3、需人工核查 5）'),
    ]),
    ('6.3 麒麟平台各组件核查脚本运行过程', None, [
        ('run_kylin_mysql', '图 10　MySQL 核查脚本（19 项：合规 6、不合规 5、需人工核查 8）'),
        ('run_kylin_redis', '图 11　Redis 核查脚本（16 项：合规 4、不合规 6、需人工核查 6）'),
        ('run_kylin_nginx', '图 12　Nginx 核查脚本（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('run_kylin_tomcat', '图 13　Tomcat 核查脚本（8 项：合规 3、需人工核查 5）'),
        ('run_kylin_dm', '图 14　达梦核查脚本（靶机无 disql 客户端，19 项标记为不适用）'),
        ('run_kylin_sqlserver', '图 15　SQL Server 核查脚本（靶机无实例，18 项标记为不适用）'),
    ]),
    ('6.4 Windows 平台核查报告', None, [
        ('fig_win_os', '图 16　操作系统核查报告（139 项：合规 28、不合规 22、需人工核查 63、不适用 26）'),
        ('fig_win_sqlserver', '图 17　SQL Server 数据库核查报告（18 项：合规 4、不合规 4、需人工核查 10）'),
        ('fig_win_net', '图 18　网络设备核查报告（23 项：合规 6、不合规 3、需人工核查 14）'),
    ]),
    ('6.5 麒麟平台核查报告', None, [
        ('fig_kylin_os', '图 19　操作系统核查报告（139 项：合规 18、不合规 9、需人工核查 84、不适用 28）'),
        ('fig_kylin_mysql', '图 20　MySQL 数据库核查报告（19 项：合规 6、不合规 5、需人工核查 8）'),
        ('fig_kylin_nginx', '图 21　Nginx 中间件核查报告（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('fig_kylin_net', '图 22　网络设备核查报告（23 项：合规 11、需人工核查 12）'),
    ]),
    ('6.6 人工核查与汇总', None, [
        ('fig_manual', '图 23　人工核查台（136 项逐项核查，支持截图取证与报告导出）'),
        ('fig_manual_filled', '图 24　人工核查台全部填写完成（136 项已核，顶部统计与各章进度同步更新）'),
        ('fig_manual_export', '图 25　人工核查台导出的 Excel 报告（136 行逐项结论，与自动脚本报告同格式）'),
        ('fig6_matrix', '图 26　《配置核查表》第 14 列「工具自动验证」逐项标注（节选，平台适用对象列从略）'),
    ]),
    ('6.7 核查项判定与实测配置对照',
     '下列截图中，每条核查项都按它在脚本源码里的原始检测命令实测一遍，输出与脚本判定放在一起对照。',
     [
        ('probe_win_1_4', '图 27　Windows 核查项 1.4 防火墙策略：判定「合规」 ↔ 实测域配置文件状态「启用」'),
        ('probe_win_1_5', '图 28　Windows 核查项 1.5 停用冗余网络设置：判定「不合规」 ↔ 实测接口 NetbiosOptions=0x0（未禁用）'),
        ('probe_win_2_11', '图 29　Windows 核查项 2.11 口令策略和屏保设置：判定「不合规」 ↔ 实测密码长度最小值 0（<8）、锁定阈值「从不」'),
        ('probe_win_1_2', '图 30　Windows 核查项 1.2 防病毒软件：判定「合规」 ↔ 实测 WMI 查询到 QI-ANXIN Tianqing 与 Windows Defender'),
        ('probe_win_2_10', '图 31　Windows 核查项 2.10 USB 接口管控：判定「不合规」 ↔ 实测 USBSTOR\\Start=0x3（未禁用，应为 4）'),
        ('probe_win_2_12', '图 32　Windows 核查项 2.12 拆除无线模块：判定「不合规」 ↔ 实测 WLAN AutoConfig 与蓝牙服务 bthserv 均为 RUNNING'),
        ('probe_win_1_3', '图 33　Windows 核查项 1.3 服务和端口裁剪：判定「需人工核查」 ↔ 实测运行中的服务清单（未发现 telnet/tftp/ftp/snmptrap 类高危服务）'),
        ('probe_kylin_1_1', '图 34　麒麟核查项 1.1 补丁安装情况：判定「不合规」 ↔ 实测 yum check-update 列出 46 个可更新软件包'),
        ('probe_kylin_1_5', '图 35　麒麟核查项 1.5 强制访问控制：判定「不合规」 ↔ 实测 SELinux status: disabled'),
        ('probe_kylin_1_6', '图 36　麒麟核查项 1.6 SSH 加密强度：判定「合规」 ↔ 实测 sshd_config 已显式配置 Ciphers 且不含弱算法'),
        ('probe_mysql_1_7', '图 37　MySQL 核查项 1.7 数据库账户管理：判定「不合规」 ↔ 实测 root 与 checker 账户 authentication_string 为空（空口令）'),
        ('probe_mysql_1_8', '图 38　MySQL 核查项 1.8 存储过程管理：判定「合规」 ↔ 实测存储过程/函数总数 2、无自定义 UDF'),
        ('probe_mysql_1_9', '图 39　MySQL 核查项 1.9 权限最小化：判定「不合规」 ↔ 实测 checker@localhost、checker@% 持有危险全局权限'),
        ('probe_mysql_1_10', '图 40　MySQL 核查项 1.10 数据库访问控制：判定「合规」 ↔ 实测仅监听 127.0.0.1:3306，未对外暴露'),
        ('probe_mysql_1_12', '图 41　MySQL 核查项 1.12 审计插件：判定「合规」 ↔ 实测 general_log=ON（已开启通用查询日志，可记录 SQL 操作）'),
        ('probe_redis_1_7', '图 42　Redis 核查项 1.7 账户管理：判定「不合规」 ↔ 实测未设置 requirepass（输出为空，无需口令即可访问）'),
        ('probe_redis_1_10', '图 43　Redis 核查项 1.10 数据库访问控制：判定「合规」 ↔ 实测 bind=127.0.0.1，访问受限'),
    ]),
]
_figno = 0
for _g, _note, _items in FIG_GROUPS:
    A(f'### {_g}')
    A('')
    if _note:
        A(_note)
        A('')
    for _fn, _cap in _items:
        _figno += 1
        A(f'![{_cap}](验证截图/{_fn}.png)')
        A('')
        _nt = FIG_NOTES.get(_fn)
        if _nt:
            A(f'> 图 {_figno}　说明：{_nt}')
            A('')

open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write(chr(10).join(L))
print(f'已生成（交付版）：{total} 项，可自动化 {n_auto}，部分可自动化 {n_part}，需人工 {n_manual}')
print('一致性校验通过：报告结论与核查表第 14 列逐项一致')
