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
doc = Document('配置核查作业指导书_v2.0.0.docx')
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

# 条目名以指导书（基准文档）为准：核查表个别条目存在错字（如"操作系应身有防火墙功能"），
# 报告中一律采用指导书标题文字，保证与基准文档一致。
for _it in guide_items:
    _c, _t = _it[1], _it[2]        # guide_items: (章, 编号, 标题)
    if _c in item_name and _t:
        item_name[_c] = re.sub(r'\s+', ' ', str(_t)).strip().replace('|', '｜')

mismatch = [c for c, lab in excel_label.items() if judge(c)[0] != lab]
if mismatch:
    print('!! 报告与核查表「工具自动验证」列不一致：', mismatch)
    sys.exit(1)

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
    'fig_kylin_redis': 'Redis 组件核查报告，报告头显示连接串（127.0.0.1:6379）与版本，16 项的判定与对应指导书条款。',
    'fig_summary_excel': 'Windows 版一键运行合并生成的汇总报告：表头标注组件数（共 8 个组件、227 项判定，其中网络设备 23 项在汇总中单列），表体按组件列出全部检查项的判定、详情、修复建议与对应指导书章节，日常上报以此文件为准。',
    'probe_mysql_1_15': '麒麟靶机 MySQL 实际环境。下半为查询 root 账户登录主机的实际输出，仅 localhost、127.0.0.1 与本机名，未开放远程登录。',
    'probe_mysql_1_19': '麒麟靶机 MySQL 实际环境。下半为查询 port 系统变量的实际输出，仍为默认端口 3306，未按要求更换。',
    'probe_mysql_1_20': '麒麟靶机 MySQL 实际环境。下半为查询 sql_mode 与 secure_file_priv 的实际输出，已启用严格模式并限定导入导出目录。',
    'probe_redis_1_11': '麒麟靶机 Redis 实际环境。下半为查询 RDB 快照规则与 AOF 开关的实际输出，save 规则已配置。',
    'probe_redis_1_19': '麒麟靶机 Redis 实际环境。下半为查询 port 的实际输出，仍为默认端口 6379。',
    'probe_redis_1_20': '麒麟靶机 Redis 实际环境。下半为查询 maxmemory 与淘汰策略的实际输出，均未配置（返回 0 与 noeviction）。',
}

FIG_GROUPS = [
    ('6.1 操作系统与网络设备核查运行过程', None, [
        ('term_win_os', 'Windows 版运行操作系统核查脚本（cscript check_xp7.vbs，139 项逐项输出判定并生成报告）'),
        ('term_win_net', 'Windows 版网络设备核查（init 生成采集清单 → 采集 → check 解析出报告，第 5 章 23 项）'),
        ('term_kylin_os', '麒麟版运行操作系统核查脚本（bash check_kylin.sh，139 项）'),
    ]),
    ('6.2 Windows 平台各组件核查脚本运行过程', None, [
        ('run_win_sqlserver', 'SQL Server 核查脚本（连接 1433 端口实跑，18 项：合规 4、不合规 4、需人工核查 10）'),
        ('run_win_mysql', 'MySQL 核查脚本（本机未部署 MySQL，19 项全部转入人工核查）'),
        ('run_win_redis', 'Redis 核查脚本（本机未部署 Redis，16 项全部转入人工核查）'),
        ('run_win_dm', '达梦核查脚本（本机无 disql 客户端，19 项标记为不适用）'),
        ('run_win_nginx', 'Nginx 核查脚本（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('run_win_tomcat', 'Tomcat 核查脚本（8 项：合规 3、需人工核查 5）'),
    ]),
    ('6.3 麒麟平台各组件核查脚本运行过程', None, [
        ('run_kylin_mysql', 'MySQL 核查脚本（19 项：合规 6、不合规 5、需人工核查 8）'),
        ('run_kylin_redis', 'Redis 核查脚本（16 项：合规 4、不合规 6、需人工核查 6）'),
        ('run_kylin_nginx', 'Nginx 核查脚本（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('run_kylin_tomcat', 'Tomcat 核查脚本（8 项：合规 3、需人工核查 5）'),
        ('run_kylin_dm', '达梦核查脚本（靶机无 disql 客户端，19 项标记为不适用）'),
        ('run_kylin_sqlserver', 'SQL Server 核查脚本（靶机无实例，18 项标记为不适用）'),
    ]),
    ('6.4 Windows 平台核查报告', None, [
        ('fig_win_os', '操作系统核查报告（139 项：合规 28、不合规 22、需人工核查 63、不适用 26）'),
        ('fig_win_sqlserver', 'SQL Server 数据库核查报告（18 项：合规 4、不合规 4、需人工核查 10）'),
        ('fig_win_net', '网络设备核查报告（23 项：合规 6、不合规 3、需人工核查 14）'),
    ]),
    ('6.5 麒麟平台核查报告', None, [
        ('fig_kylin_os', '操作系统核查报告（139 项：合规 18、不合规 9、需人工核查 84、不适用 28）'),
        ('fig_kylin_mysql', 'MySQL 数据库核查报告（19 项：合规 6、不合规 5、需人工核查 8）'),
        ('fig_kylin_nginx', 'Nginx 中间件核查报告（8 项：合规 2、需人工核查 5、不适用 1）'),
        ('fig_kylin_redis', 'Redis 数据库核查报告（16 项：合规 4、不合规 6、需人工核查 6）'),
        ('fig_kylin_net', '网络设备核查报告（23 项：合规 11、需人工核查 12）'),
    ]),
    ('6.6 人工核查与汇总', None, [
        ('fig_manual', '人工核查台（136 项逐项核查，支持截图取证与报告导出）'),
        ('fig_manual_filled', '人工核查台全部填写完成（136 项已核，顶部统计与各章进度同步更新）'),
        ('fig_manual_export', '人工核查台导出的 Excel 报告（136 行逐项结论，与自动脚本报告同格式）'),
        ('fig_summary_excel', '汇总报告 Excel 视图（8 个组件、227 项判定，上报用）'),
        ('fig6_matrix', '《配置核查表》第 14 列「工具自动验证」逐项标注（节选，平台适用对象列从略）'),
    ]),
    ('6.7 核查项判定与实测配置对照',
     None, [
        ('probe_win_1_4', 'Windows 核查项 1.4 防火墙策略：判定「合规」 ↔ 实测域配置文件状态「启用」'),
        ('probe_win_1_5', 'Windows 核查项 1.5 停用冗余网络设置：判定「不合规」 ↔ 实测接口 NetbiosOptions=0x0（未禁用）'),
        ('probe_win_2_11', 'Windows 核查项 2.11 口令策略和屏保设置：判定「不合规」 ↔ 实测密码长度最小值 0（<8）、锁定阈值「从不」'),
        ('probe_win_1_2', 'Windows 核查项 1.2 防病毒软件：判定「合规」 ↔ 实测 WMI 查询到 QI-ANXIN Tianqing 与 Windows Defender'),
        ('probe_win_2_10', 'Windows 核查项 2.10 USB 接口管控：判定「不合规」 ↔ 实测 USBSTOR\\Start=0x3（未禁用，应为 4）'),
        ('probe_win_2_12', 'Windows 核查项 2.12 拆除无线模块：判定「不合规」 ↔ 实测 WLAN AutoConfig 与蓝牙服务 bthserv 均为 RUNNING'),
        ('probe_kylin_1_1', '麒麟核查项 1.1 补丁安装情况：判定「不合规」 ↔ 实测 yum check-update 列出 46 个可更新软件包'),
        ('probe_kylin_1_5', '麒麟核查项 1.5 强制访问控制：判定「不合规」 ↔ 实测 SELinux status: disabled'),
        ('probe_kylin_1_6', '麒麟核查项 1.6 SSH 加密强度：判定「合规」 ↔ 实测 sshd_config 已显式配置 Ciphers 且不含弱算法'),
        ('probe_mysql_1_7', 'MySQL 核查项 1.7 数据库账户管理：判定「不合规」 ↔ 实测 root 与 checker 账户 authentication_string 为空（空口令）'),
        ('probe_mysql_1_8', 'MySQL 核查项 1.8 存储过程管理：判定「合规」 ↔ 实测存储过程/函数总数 2、无自定义 UDF'),
        ('probe_mysql_1_9', 'MySQL 核查项 1.9 权限最小化：判定「不合规」 ↔ 实测 checker@localhost、checker@% 持有危险全局权限'),
        ('probe_mysql_1_10', 'MySQL 核查项 1.10 数据库访问控制：判定「合规」 ↔ 实测仅监听 127.0.0.1:3306，未对外暴露'),
        ('probe_mysql_1_12', 'MySQL 核查项 1.12 审计插件：判定「合规」 ↔ 实测 general_log=ON（已开启通用查询日志，可记录 SQL 操作）'),
        ('probe_redis_1_7', 'Redis 核查项 1.7 账户管理：判定「不合规」 ↔ 实测未设置 requirepass（输出为空，无需口令即可访问）'),
        ('probe_redis_1_10', 'Redis 核查项 1.10 数据库访问控制：判定「合规」 ↔ 实测 bind=127.0.0.1，访问受限'),
        ('probe_mysql_1_15', 'MySQL 核查项 1.15 远程访问控制：判定「合规」 ↔ 实测 root 账户仅允许本机登录'),
        ('probe_mysql_1_19', 'MySQL 核查项 1.19 更换默认端口：判定「不合规」 ↔ 实测仍使用默认端口 3306'),
        ('probe_mysql_1_20', 'MySQL 核查项 1.20 数据库安全策略：判定「合规」 ↔ 实测 sql_mode 严格模式、secure_file_priv 已限定'),
        ('probe_redis_1_11', 'Redis 核查项 1.11 备份策略：判定「合规」 ↔ 实测已配置 RDB 快照（save 3600/300/60）'),
        ('probe_redis_1_19', 'Redis 核查项 1.19 更换默认端口：判定「不合规」 ↔ 实测仍使用默认端口 6379 且未设口令'),
        ('probe_redis_1_20', 'Redis 核查项 1.20 数据库安全策略：判定「不合规」 ↔ 实测 maxmemory 与淘汰策略均未配置'),
    ]),
]

# ---------- 7. 生成报告（交付版：结论 + 方法 + 覆盖统计 + 项号索引）----------
L = []
A = L.append

def idx(codes):
    return '、'.join(codes) if codes else '—'

_T = [0]
TNO = {}


def _evidence_cell(code, st, reason):
    """逐项「对应证据」：有抽查图给图号；工具核查项给脚本名；人工项标人工核查。"""
    _ef = ITEM_FIGS.get(code)
    if _ef:
        return '、'.join('图 %d' % _n for _n in sorted(_ef))
    if st == '需人工':
        return '人工核查'
    _sc = re.findall(r'check_[A-Za-z0-9_]+\.(?:vbs|sh|ps1)', reason or '')
    _seen, _uniq = set(), []
    for _x in _sc:
        if _x not in _seen:
            _seen.add(_x); _uniq.append(_x)
    if _uniq:
        return '、'.join(_uniq[:3])
    _parts = [p.strip() for p in re.split(r'[；;]', reason or '') if p.strip()]
    if _parts:
        _txt = '；'.join(_parts[:2]).rstrip('；;、,，')
        _txt = _txt[:(26 if len(_parts) > 2 else 30)].rstrip('；;、,，')
        if _txt.count('（') > _txt.count('）'):        # 截断留下半个括号时，砍到括号前
            _txt = _txt[: _txt.rfind('（')].rstrip('；;、,，')
        return _txt + ('等' if len(_parts) > 2 else '')
    return '脚本核查'


def _tab(name, key=None):
    _T[0] += 1
    A(f'表 {_T[0]}　{name}')
    A('')
    if key:
        TNO[key] = _T[0]
    return _T[0]


def by_state(ch, st):
    return [c for c, _t, _a, s2, _r in detail[ch] if s2 == st]

A('%# 西北工业大学制')
A('% 配置核查指导书')
A('% 和工具交叉验证报告')
A('%> 验证基准：《配置核查作业指导书》v2.0.0（10 章 ' + str(total) + ' 项）')
A('%> 验证对象：配置核查工具（脚本、网络设备核查、人工核查台）')
A('%$ 2026 年 9 月')
A('')
A('本报告对《配置核查作业指导书》v2.0.0 所列检查项的自动化核查能力进行交叉验证，'
  '目的是确认指导书要求的安全配置检查项中，哪些可由工具直接给出判定、哪些需结合台账或现场情况确认、'
  '哪些必须由人工完成，为核查作业的组织与结果采信提供依据。'
  '验证对象为配置核查工具（以下简称工具），包含 Windows 版、麒麟版共 16 个核查脚本及人工核查台；'
  '核查对象涵盖操作系统、数据库、中间件与网络设备四类。'
  '验证方法为逐条比对核查脚本源码中的检测分支与输出结论，并对抽查项以原始检测命令实测复核，'
  '全过程证据随文附列。')
A('')
A('| 文档信息 | 内容 |')
A('|:--|:--|')
A('| 文档版本 | v2.0.0 |')
A(f'| 报告日期 | {today} |')
A('| 基准文档 | 《配置核查作业指导书》v2.0.0（10 章、' + str(total) + ' 个编号检查项） |')
A('| 验证对象 | 配置核查工具（Windows 版、麒麟版、人工核查台） |')
A('| 结果口径 | 与《配置核查表》第 14 列「工具自动验证」逐项一致 |')
A('')
A('| 版本 | 日期 | 修订说明 |')
A('|:--:|:--:|:--|')
A(f'| v2.0.0 | {today} | 修订发布：统一文档版本号与编号体系，图表补标准标题并在正文引用，'
  '验证证据按核查栏目就近对照，首段补充测试目的、对象与方法。 |')
A('')
A('# 1 验证结论')
A('')
A(f'本次验证以《配置核查作业指导书》v2.0.0（10 章、{total} 个编号检查项）为基准，'
  f'逐条比对工具的核查能力，三类验证结论的项数与占比见表 1；'
  f'各检查项的具体结论见第 4 章，支撑证据随各节附列。')
A('')
_tab('验证结论统计', 'concl')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 纳入工具核查 | {n_auto + n_part} | {(n_auto + n_part)/total*100:.1f}% |')
A(f'| 　其中：可自动化（工具直接给出合格/不合格判定） | {n_auto} | {n_auto/total*100:.1f}% |')
A(f'| 　其中：部分可自动化（脚本已核查，结论需结合台账或现场确认） | {n_part} | {n_part/total*100:.1f}% |')
A(f'| 人工核查 | {n_manual} | {n_manual/total*100:.1f}% |')
A('')
A(f'全部 {total} 个检查项中，{n_auto + n_part} 项可由工具执行：其中 {n_auto} 项工具直接给出合格或不合格判定（占 {n_auto/total*100:.1f}%），'
  f'{n_part} 项脚本完成检测后仍需结合设备台账或现场情况确认结论（占 {n_part/total*100:.1f}%）。'
  f'其余 {n_manual} 项（占 {n_manual/total*100:.1f}%）属现场查看、文档调阅或实测验证类，'
  f'按标准要求由人工完成核查并留存证据，不属于工具能力缺陷。'
  f'阅读本报告时需注意：{n_auto + n_part} 项表示工具可执行核查的范围，'
  f'能直接给出判定的为 {n_auto} 项，二者含义不同。')
A('')
A('# 2 验证对象与方法')
A('')
A('## 2.1 验证对象')
A('')
A('（1）指导书：《配置核查作业指导书》v2.0.0，共 10 章、' + str(total) + ' 个编号检查项。')
A('（2）工具：16 个脚本。操作系统核查 2 个（Windows、麒麟各 1）；数据库与中间件核查 12 个（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat，两平台各 6）；网络设备核查 2 个（麒麟版与 Windows 版，采集-解析模式，判定逻辑与采集文件通用，覆盖第 5 章 23 项）。')
A('（3）项数口径：操作系统核查脚本输出 139 项判定，其中 136 项与指导书编号检查项一一对应，另 3 项（口令双因子认证、日志留存完整性、密钥签发与证书管控）为脚本在编号体系外自行补充的核查条目；本报告统计一律以指导书 136 个编号检查项为基准。')
A('（4）工具的组成、部署方式与使用方法详见《配置核查工具操作说明书》。')
A('')
A('## 2.2 验证方法')
A('')
A('从指导书提取全部编号检查项，先确定每项的适用对象（操作系统、数据库、中间件、网络设备），再按对象核对脚本的核查能力。'
  '判定依据是脚本源码里的检测分支和结果输出：能给出合格或不合格判定的列为可自动化；'
  '执行了检测命令、结论还需结合台账或现场确认的列为部分可自动化；'
  '其余属现场查看、文档调阅或实测验证类的列为需人工核查。')
A('')
A('## 2.3 判定规则')
A('')
A('三类验证结论的判定条件见表 2。')
A('')
_tab('验证结论判定条件', 'rule')
A('| 验证结论 | 判定条件 |')
A('|---|---|')
A('| 可自动化 | 该检查项适用的全部对象，脚本均可自动给出合格或不合格判定 |')
A('| 部分可自动化 | 脚本已执行核查，受核查对象或结论性质所限，需结合设备台账或现场情况确认最终结论 |')
A('| 需人工核查 | 现场查看、文档调阅、实测验证类检查项，由人工按指导书方法完成并留存证据 |')
A('')
A('主机防火墙类检查项（第 1.4 项）的判定口径：工具按 firewalld、ufw、iptables 的顺序识别主机防火墙，'
  '任一种生效即视为具备防火墙功能。其中 firewalld 与 ufw 是 netfilter 的管理前端，iptables 即 netfilter 本体；'
  '麒麟类系统在等级保护加固中常见「关闭 firewalld、以固化 iptables 规则集」的部署方式，'
  '因此脚本在 firewalld/ufw 均未运行、但 iptables 已配置规则集时判定为合规，'
  '判定依据为规则集存在本身；规则集是否随重启持久化、默认策略是否非全放行，属该分支的判定边界，'
  '需结合现场情况确认（见 2.4）。')
A('')
A('## 2.4 验证范围与局限')
A('')
A('本次验证的实测环境与覆盖范围见表 3，据此界定结论的适用范围。')
A('')
_tab('实测环境与覆盖范围', 'scope')
A('| 对象 | 实测环境 | 本次验证情况 |')
A('|---|---|---|')
A('| 操作系统 | Windows 10 专业版（本机） | 脚本实跑，139 项判定全部产出 |')
A('| 操作系统 | 银河麒麟 V10 靶机 | 脚本实跑，139 项判定全部产出 |')
A('| 数据库 | MySQL / MariaDB 10.3.39（麒麟靶机） | 脚本实跑，19 项判定全部产出 |')
A('| 数据库 | Redis（麒麟靶机） | 脚本实跑，16 项判定全部产出 |')
A('| 数据库 | SQL Server 16.0.4265.3（Windows 本机） | 脚本实跑，18 项判定全部产出 |')
A('| 数据库 | 达梦 DM8 | 环境缺 disql 客户端，19 项按不适用处理，未实测 |')
A('| 中间件 | Nginx、Tomcat（两平台） | 脚本实跑，各 8 项判定全部产出 |')
A('| 网络设备 | 华为、华三、锐捷回显样例 | 按采集-解析方式执行，23 项判定产出 |')
A('| 操作系统 | Windows 7 / Windows XP / 中标麒麟 | 脚本声明兼容，本次未在实机验证 |')
A('')
A('结论的局限如下：')
A('')
A('（1）判定依据为脚本源码中的检测分支与输出结论，逐项比对得出，可说明工具「具备」该项检测能力；'
  '判定逻辑本身是否正确，由 17 项抽查项的原始检测命令实测复核覆盖（占全部检查项的 12.5%，集中于第 1、2 章）。')
A('（2）抽查项为脚本检测命令的独立复现，未引入未参与开发的第三方独立复评，'
  '因此本报告给出的是工具自身能力与实测一致性的验证结果，不构成第三方测评结论。')
A('（3）网络设备 23 项的结论来自运维人员登录设备采集的回显文件，回显的真实性与完整性由采集方负责。')
A('（4）未实测平台（Windows 7 / XP、中标麒麟）的结论为脚本声明兼容，需在对应实机复核后采信。')
A('（5）操作系统脚本输出 139 项判定，其中 136 项对应指导书编号检查项，另 3 项为脚本在编号体系外补充的核查条目；'
  '汇总报告 227 项为 7 个自动核查组件判定之和，网络设备 23 项在汇总中单列。')
A('（6）主机防火墙类条目（1.4）在 firewalld/ufw 未运行、依赖 iptables 规则集判合规的场景，'
  '脚本不校验规则集的持久化与默认策略；此类系统的最终结论建议结合现场确认规则是否随重启生效。')
A('')
A('# 3 各章覆盖统计')
A('')
A('各章检查项的验证结论分布见表 3，据此可判断自动化核查能力在各章之间的差异。')
A('')
_tab('各章覆盖统计', 'cover')
A('| 章节 | 检查项 | 可自动化 | 部分可自动化 | 需人工核查 |')
A('|---|---|---|---|---|')
for ch in chapters:
    a, b, c = stat[ch]
    A(f'| {ch} | {len(detail[ch])} | {a} | {b} | {c} |')
A('')
A('# 4 逐项验证结果')
A('')
A('以下按《配置核查作业指导书》章节顺序逐项列示验证结论，本章小节号 4.1～4.10 对应指导书第 1～10 章；'
  '每项的核查对象、所用脚本与结论依据见《配置核查表》第 14 列「工具自动验证」。')
A('')
# 图件归属（0 = 工具整体）与图号预分配，供正文引用
FIG_CHAPTER = {}
for _fn in ('term_win_os', 'term_kylin_os', 'fig_win_os', 'fig_kylin_os',
            'fig_manual', 'fig_manual_filled', 'fig_manual_export',
            'fig_summary_excel', 'fig6_matrix'):
    FIG_CHAPTER[_fn] = 0
for _fn in ('term_win_net', 'fig_win_net', 'fig_kylin_net'):
    FIG_CHAPTER[_fn] = 5
for _fn in ('run_win_sqlserver', 'run_win_mysql', 'run_win_redis', 'run_win_dm',
            'run_win_nginx', 'run_win_tomcat', 'run_kylin_mysql', 'run_kylin_redis',
            'run_kylin_nginx', 'run_kylin_tomcat', 'run_kylin_dm', 'run_kylin_sqlserver',
            'fig_win_sqlserver', 'fig_kylin_mysql', 'fig_kylin_nginx', 'fig_kylin_redis'):
    FIG_CHAPTER[_fn] = 1
FIG_CAP = {}
for _g, _note, _items in FIG_GROUPS:
    for _fn, _cap in _items:
        FIG_CAP[_fn] = _cap
        if _fn.startswith('probe_'):
            FIG_CHAPTER[_fn] = int(re.match(r'probe_\w+_(\d+)_', _fn).group(1))
FIG_NO, _fc = {}, 0
for _ch in list(range(1, 11)) + [0]:
    for _g, _note, _items in FIG_GROUPS:
        for _fn, _cap in _items:
            if FIG_CHAPTER.get(_fn) == _ch:
                _fc += 1
                FIG_NO[_fn] = _fc
ITEM_FIGS = {}
_no_tool_evid = 0
for _fn, _no in FIG_NO.items():
    if _fn.startswith('probe_'):
        _m = re.match(r'probe_\w+_(\d+)_(\d+)$', _fn)
        ITEM_FIGS.setdefault('%s.%s' % (_m.group(1), _m.group(2)), []).append(_no)

for _ci, ch in enumerate(chapters, 1):
    _cname = ch.split(" ", 1)[1]
    A(f'## 4.{_ci} {_cname}')
    A('')
    _tno = _T[0] + 1
    A(f'本章共 {len(detail[ch])} 项检查项，逐项验证结论与对应证据见表 {_tno}。')
    A('')
    _tab(f'{_cname}检查项逐项验证结果')
    A('| 编号 | 检查项 | 验证结论 | 对应证据 |')
    A('|---|---|---|---|')
    for code, title, appl, st, reason in detail[ch]:
        A(f'| {code} | {item_name.get(code) or title} | {st} | {_evidence_cell(code, st, reason)} |')
    A('')
    _figs = [(FIG_NO[_fn], _fn, _cap) for _g2, _n2, _it in FIG_GROUPS for _fn, _cap in _it
             if FIG_CHAPTER.get(_fn) == _ci]
    _figs.sort()
    _n_all = len(detail[ch])
    _n_auto = sum(1 for _c, _t, _a, _s, _r in detail[ch] if _s == '可自动化')
    _n_part = sum(1 for _c, _t, _a, _s, _r in detail[ch] if _s == '部分可自动化')
    _n_tool = _n_auto + _n_part
    _n_man = _n_all - _n_tool
    if _figs:
        _lo, _hi = _figs[0][0], _figs[-1][0]
        A(f'本章 {_n_all} 项中，{_n_tool} 项由工具执行（{_n_auto} 项工具可直接给出判定，'
          f'{_n_part} 项脚本检测后需结合台账或现场确认），{_n_man} 项由人工核查；'
          f'工具核查证据见图 {_lo}～图 {_hi}，取自工具在实际环境运行的输出与原始检测命令实测结果。')
        A('')
        for _no, _fn, _cap in _figs:
            A(f'![图 {_no}　{_cap}](验证截图/{_fn}.png)')
            A('')
            _nt = FIG_NOTES.get(_fn)
            if _nt:
                A(f'> 附图 {_no}　{_nt}')
                A('')
    elif _n_tool:
        _det = (f'{_n_auto} 项工具可直接给出判定' + (f'，{_n_part} 项脚本检测后需结合台账或现场确认' if _n_part else ''))
        _tail = (f'其余 {_n_man} 项属现场查看、文档调阅与实测验证类，由人工按《配置核查作业指导书》的方法逐项核查并留存证据。'
                 if _n_man else '')
        A(f'本章 {_n_all} 项中，{_n_tool} 项由工具执行（{_det}），'
          f'所用脚本与检测项见上表「对应证据」列，完整标注见《配置核查表》第 14 列；{_tail}')
        A('')
    else:
        _no_tool_evid += 1
        A('本章检查项属现场查看、文档调阅与实测验证类，工具不产出具机证据，'
          '由人工按《配置核查作业指导书》的方法逐项核查并留存证据；'
          '工具对本章的验证结论以《配置核查表》第 14 列逐项标注为准。'
          if _no_tool_evid == 1 else
          '本章检查项同属现场查看与文档调阅类，工具不产出具机证据，由人工按指导书方法核查并留存证据。')
        A('')
A('# 5 说明')
A('')
_tno5 = _T[0] + 1
A(f'本次验证需说明的事项见表 {_tno5}。')
A('')
_tab('验证说明事项', 'note')
A('| 事项 | 说明 |')
A('|---|---|')
A('| 人工核查范围 | 第 6 至 9 章（物理安全、组织机构、规章制度、管理实施）为现场查看与文档调阅类；第 1 至 5 章中涉及应用代码安全、渗透验证、业务架构与台账比对的检查项，按标准要求由人工完成核查。 |')
A('| 网络设备核查方式 | 第 5 章按工具生成的命令清单，由运维人员登录交换机、防火墙等设备执行并留存回显，工具据此逐条判定并出具报告。 |')
A('| 结果合并 | 工具与人工核查台采用同一套检查项编号与判定规则，两部分结果可直接合并汇总。 |')
A('')
# 第 6 章引导句在 FIG_GROUPS 定义后统一生成（需累计图号）
# 每张图的说明（环境、方法、关键数据），排在图片下方，便于甲方看图即懂
GROUP_LEADS = {
    '6.1 操作系统与网络设备核查运行过程':
        '两平台操作系统核查脚本的完整运行过程，以及网络设备「采集—解析」两阶段执行实况（图 {n0}～图 {n1}）。',
    '6.2 Windows 平台各组件核查脚本运行过程':
        'Windows 平台 6 个数据库与中间件核查脚本逐个实跑（图 {n0}～图 {n1}）：已部署组件给出逐项判定，未部署组件按降级策略照常出具报告。',
    '6.3 麒麟平台各组件核查脚本运行过程':
        '麒麟平台 6 个数据库与中间件核查脚本在靶机上逐个实跑（图 {n0}～图 {n1}）。',
    '6.4 Windows 平台核查报告':
        'Windows 平台生成的核查报告样例：操作系统、SQL Server 与网络设备（图 {n0}～图 {n1}）。',
    '6.5 麒麟平台核查报告':
        '麒麟平台生成的核查报告样例（图 {n0}～图 {n1}）：操作系统与各数据库、中间件组件。',
    '6.6 人工核查与汇总':
        '人工核查台填写与导出、自动汇总报告及《配置核查表》标注情况（图 {n0}～图 {n1}）。',
    '6.7 核查项判定与实测配置对照':
        '抽查各对象典型检查项（图 {n0}～图 {n1}）：每条均按脚本源码中的原始检测命令实测一遍，脚本判定与实测输出并排对照。',
}

A('# 6 工具整体运行与结果汇总证据')
A('')
A('第 4 章各节已就近列出该章检查项的证据；本节补充工具整体运行、报告输出与人工核查环节的证据，'
  '以反映工具的实际运行情况，不与具体检查项一一对应。')
A('')

CH0_GROUPS = [
    ('6.1 操作系统核查运行过程与报告',
     '两平台操作系统核查脚本的实际运行画面与生成的核查报告（图 {n0}～图 {n1}）。'
     '该脚本覆盖第 1～10 章中可在操作系统层核查的检查项。',
     ['term_win_os', 'term_kylin_os', 'fig_win_os', 'fig_kylin_os']),
    ('6.2 人工核查与结果汇总',
     '人工核查台的填写、导出与自动汇总报告（图 {n0}～图 {n1}），'
     '以及《配置核查表》第 14 列「工具自动验证」的逐项标注情况。',
     ['fig_manual', 'fig_manual_filled', 'fig_manual_export', 'fig_summary_excel', 'fig6_matrix']),
]
for _gname, _lead, _fns in CH0_GROUPS:
    _nos = sorted(FIG_NO[_f] for _f in _fns)
    A(f'## {_gname}')
    A('')
    A(_lead.format(n0=_nos[0], n1=_nos[-1]))
    A('')
    for _fn in _fns:
        A(f'![图 {FIG_NO[_fn]}　{FIG_CAP[_fn]}](验证截图/{_fn}.png)')
        A('')
        _nt = FIG_NOTES.get(_fn)
        if _nt:
            A(f'> 附图 {FIG_NO[_fn]}　{_nt}')
            A('')

open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write(chr(10).join(L))
print(f'已生成（交付版）：{total} 项，可自动化 {n_auto}，部分可自动化 {n_part}，需人工 {n_manual}')
print('一致性校验通过：报告结论与核查表第 14 列逐项一致')
