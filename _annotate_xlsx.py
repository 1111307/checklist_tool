# -*- coding: utf-8 -*-
"""核查表第13列「工具自动验证」详情标注 v3。

判定标准（用户定义）：脚本是否实际执行核查。
  auto    = 脚本有自动判定分支（pass/fail）
  check   = manual 但脚本实际执行了检测（详情含检测输出/变量插值/采集结果，结论留人工确认）
  reminder= manual 且纯文字提示（脚本未执行任何检测，仅引用指导书方法）
分类：
  可自动化     = 全部适用对象（勾选矩阵）的脚本均可自动判定
  部分可自动化 = 适用对象部分可自动 / 脚本采集了证据但结论需人工（第5章含部分信号项）
  需人工       = 适用对象的脚本均未执行检测（纯提示），或第6-9章现场类
适用对象以核查表 C~L 勾选矩阵为准（√=适用、—=不适用）。
用法：python _annotate_xlsx.py [--apply]
"""
import re, glob, os, sys, math
from collections import defaultdict, Counter
from openpyxl import load_workbook
from openpyxl.styles import Alignment

XLSX = '配置核查表_v2.0.0_标注自动验证.xlsx'
APPLY = '--apply' in sys.argv
BSLASH = chr(92)

# ---------- 1. 脚本提取：状态 + manual 证据 ----------
ST_SH = re.compile(r'add_result\s+"([^"]+)"\s+"[^"]*"\s+"[^"]*"\s+"([^"]+)"')
ST_VBS = re.compile(r'AddResult\s*\(?\s*"([^"]+)"\s*,\s*"[^"]*"\s*,\s*"[^"]*"\s*,\s*"([^"]+)"')
EVIDENCE = re.compile(
    r'\$\{?\w|\$\(|"\s*&|检测到|未检测到|采集到|已安装|已部署|已启用|未发现|已划分|'
    r'执行失败|无法读取|无法查询|无法确认|无法获取|无法确定|权限不足|已建立|'
    r'未找到|未运行|未在|（输出|输出:')
# 「无法连接 X（…），请检查 XX_PASS/环境变量」是连接探测失败提示，不是对该条目的检测结果；
# 若剥掉变量插值与连接告警后没有实际检测叙述，则不能算作"脚本已执行该项检测"
CONN_FAIL = re.compile(r'无法连接|连接失败|连接超时|不可达|请检查\s*\S*(PASS|USER|HOST|PORT)|环境变量')
DETECT_WORD = re.compile(r'检测到|采集到|已安装|已启用|已划分|未发现|发现|总数|条目|个可|条可')

def has_evidence(det, concat):
    if not (EVIDENCE.search(det) or concat):
        return False
    if CONN_FAIL.search(det):
        stripped = re.sub(r'\$\{?\w[\w}]*|\$\([^)]*\)', '', det)
        stripped = CONN_FAIL.sub('', stripped)
        return bool(DETECT_WORD.search(stripped))
    return True

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
    if not is_vbs:
        # bash 里把编号当参数传给包装函数时（如 db_na_or_manual "1.7" ...），
        # 真正落盘的是函数内的 add_result "$1"，直接扫会漏掉这些条目的能力
        for fm in re.finditer(r'(?m)^([A-Za-z_]\w*)\s*\(\s*\)\s*\{', data):
            name = fm.group(1)
            end = data.find('\n}', fm.end())
            body = data[fm.end():end if end > 0 else len(data)]
            if 'add_result "$1"' not in body:
                continue
            wstat, wman = set(), []
            for m in st.finditer(body):
                if m.group(1) != '$1':
                    continue
                wstat.add(m.group(2))
                if m.group(2) == 'manual':
                    det, concat = read_detail(body, m.end(), False)
                    wman.append((det, concat))
            if not wstat:
                continue
            for cm in re.finditer(r'\b%s\s+"([0-9]+\.[0-9]+[^"]*)"' % re.escape(name), data):
                code = cm.group(1)
                d[code]['statuses'] |= wstat
                d[code]['manuals'].extend(wman)
    return d

def script_cap(info):
    """脚本对单项的能力：auto > check > reminder > na > none"""
    kinds = set(info['statuses']) - {'na'}
    if 'pass' in kinds or 'fail' in kinds:
        return 'auto'
    if 'manual' in kinds:
        for det, concat in info['manuals']:
            if has_evidence(det, concat):
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

# ---------- 2. 适用矩阵 ----------
wb = load_workbook(XLSX)
ws = wb.active
COLS = {3: 'Windows', 4: 'Windows', 5: '麒麟', 6: '麒麟', 7: 'Nginx', 8: 'Tomcat',
        9: '网络设备', 10: 'MySQL', 11: 'SQL Server', 12: '达梦', 13: 'Redis'}
TARGET_SCRIPTS = {
    'Windows': ['win/check_xp7'], '麒麟': ['kylin/check_kylin'],
    'Nginx': ['win/check_nginx', 'kylin/check_nginx'], 'Tomcat': ['win/check_tomcat', 'kylin/check_tomcat'],
    '网络设备': ['kylin/check_network'],
    'MySQL': ['win/check_mysql', 'kylin/check_mysql'],
    'SQL Server': ['win/check_sqlserver', 'kylin/check_sqlserver'],
    '达梦': ['win/check_dm', 'kylin/check_dm'], 'Redis': ['win/check_redis', 'kylin/check_redis'],
}
# 标注文案里展示的脚本名
SCRIPT_NAME = {
    'Windows': 'check_xp7.vbs', '麒麟': 'check_kylin.sh',
    'Nginx': 'check_nginx', 'Tomcat': 'check_tomcat',
    '网络设备': 'check_network.sh / check_network.ps1',
    'MySQL': 'check_mysql', 'SQL Server': 'check_sqlserver',
    '达梦': 'check_dm', 'Redis': 'check_redis',
}
TYPE_ORDER = ['Windows', '麒麟', 'Nginx', 'Tomcat', '网络设备', 'MySQL', 'SQL Server', '达梦', 'Redis']
# 文案中可接受的指代（脚本名/对象名/所属类别），用于校验「部分可自动化」行是否交代清楚
FAMILY = {
    'Windows': ['check_xp7', 'Windows', 'windows'],
    '麒麟': ['check_kylin', '麒麟'],
    'Nginx': ['check_nginx', 'Nginx', '中间件'],
    'Tomcat': ['check_tomcat', 'Tomcat', '中间件'],
    '网络设备': ['check_network', '网络设备', '边界设备', '设备回显'],
    'MySQL': ['check_mysql', 'MySQL', '数据库'],
    'SQL Server': ['check_sqlserver', 'SQL Server', '数据库'],
    '达梦': ['check_dm', '达梦', '数据库'],
    'Redis': ['check_redis', 'Redis', '数据库'],
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

# ---------- 3. 文案 ----------
# 全自动特例：写明具体检测内容
HAND_AUTO = {
    '2.3':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查自动登录与空口令账户限制）',
    '2.7':  '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查系统盘文件保护状态）',
    '5.23': '可自动化（用 check_network.sh / check_network.ps1 采集端口状态并计算闲置率自动判定；配备必要性可结合台账复核）',
    '10.1': '可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查 TLS/SSH 协议版本；协议层条目，勾选矩阵未单列对象）',
}
# 部分可自动化（脚本已执行核查但覆盖/结论不完全）
HAND_PARTIAL = {
    '1.1':  '部分可自动化（用 check_xp7.vbs、check_kylin.sh 及 MySQL、Redis、SQL Server、Nginx、Tomcat 组件脚本自动核查；达梦脚本采集版本信息、补丁比对需人工）',
    '1.8':  '部分可自动化（用 check_xp7.vbs 及 MySQL、SQL Server、Redis 组件脚本自动核查；check_dm 采集存储过程数量，check_kylin.sh 检测数据库服务运行状态，账户与冗余判定需人工核查）',
    '1.13': '部分可自动化（用 check_xp7.vbs 自动核查多库混用情况；check_kylin.sh 检测数据库服务运行状态，MySQL、SQL Server、达梦侧脚本未执行该项检测，分类独立存储需人工核查）',
    '1.16': '部分可自动化（用 check_xp7.vbs 检测 MySQL 严格 SQL 模式自动判定；麒麟及各数据库脚本未执行该项检测，应用层参数校验需人工核查）',
    '1.20': '部分可自动化（用 check_xp7.vbs 及 MySQL、达梦、Redis 组件脚本自动核查；check_kylin.sh 检测数据库服务运行状态，SQL Server 侧脚本未执行该项检测，安全策略明细需人工核查）',
    '1.21': '部分可自动化（用 check_xp7.vbs 及 MySQL、SQL Server 组件脚本自动核查；check_kylin.sh 检测数据库服务运行状态，达梦、Redis 侧审计粒度需人工核查）',
    '1.22': '部分可自动化（用 check_xp7.vbs 自动核查本机 MySQL binlog 日志保留天数；check_kylin.sh 检测数据库服务运行状态，SQL Server、达梦、Redis 侧脚本未执行该项检测，留存时长与独立监控需人工核查）',
    '1.24': '部分可自动化（用 check_xp7.vbs、check_kylin.sh 及 MySQL 组件脚本自动核查；SQL Server、达梦、Redis 侧留存时长需人工核查）',
    '1.25': '部分可自动化（用 check_xp7.vbs 及 MySQL、SQL Server、达梦、Redis 组件脚本自动核查；麒麟侧需人工）',
    '2.2':  '部分可自动化（check_xp7.vbs 自动判定；麒麟侧以 1.1 补丁核查结果为准、结合终端管理平台人工核实）',
    '2.5':  '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本统计自启服务数、多余服务判定需人工）',
    '2.6':  '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本统计网络连接数、违规外联判定需人工）',
    '2.8':  '部分可自动化（check_xp7.vbs 自动判定；麒麟侧需人工确认终端管控软件部署）',
    '2.15': '部分可自动化（check_xp7.vbs 自动判定；麒麟侧需结合终端管理平台人工核查）',
    '2.16': '部分可自动化（用 check_xp7.vbs、check_kylin.sh 及 MySQL、SQL Server、Redis 组件脚本自动核查；达梦脚本采集版本、升级比对需人工）',
    '3.1':  '部分可自动化（用 check_xp7.vbs 自动核查 BitLocker、check_kylin.sh 自动核查 LUKS；数据库层 TDE/列级加密未执行该项检测，需人工核查）',
    '3.5':  '部分可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查主机日志保护；中间件、数据库与网络设备侧日志加密传输需人工核查）',
    '3.6':  '部分可自动化（用 check_xp7.vbs、check_kylin.sh 自动核查终端 DLP 客户端；边界设备内容过滤与敏感内容识别未执行该项检测，需人工核查）',
    '3.8':  '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本检测分区划分、分级授权需人工核查；数据库侧脚本未执行该项检测）',
    '3.10': '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本检测共享服务、目录权限需人工核查）',
    '4.1':  '部分可自动化（check_kylin.sh 检测备份配置与定时任务、备份有效性需人工确认；Windows 侧需人工；数据库侧脚本未执行该项检测）',
    '4.5':  '部分可自动化（用 check_nginx 自动核查限速配置；check_xp7.vbs、check_kylin.sh、check_tomcat 仅输出方法提示，边界设备 DDoS 清洗需人工核查）',
    '4.8':  '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本检测防篡改工具、Web 目录监控配置需人工确认）',
    '4.11': '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本列出管理员组、RBAC 细粒度需人工核查；数据库侧脚本未执行该项检测）',
    '4.12': '部分可自动化（check_xp7.vbs 自动判定；麒麟侧需按指导书方法人工核查并发限制；中间件（Nginx/Tomcat）脚本未执行该项检测）',
    '4.17': '部分可自动化（check_xp7.vbs 自动判定；麒麟脚本统计 sudo 配置、三权分立划分需人工核查；数据库侧脚本未执行该项检测）',
    '4.23': '部分可自动化（用 check_xp7.vbs、check_tomcat 自动核查；麒麟侧需人工核查管理端口与业务端口分离；数据库侧脚本未执行该项检测）',
    '4.24': '部分可自动化（check_xp7.vbs 自动判定；麒麟侧需人工核查应用与数据分离部署；数据库脚本未执行该项检测）',
    '4.31': '部分可自动化（check_kylin.sh 自动判定；Windows 侧需人工核查并发会话限制；Nginx 脚本未执行该项检测）',
    '4.32': '部分可自动化（用 check_xp7.vbs 自动核查备份配置；check_kylin.sh 检测备份目录/备份文件/定时任务、恢复能力需人工验证；数据库侧脚本未执行该项检测）',
    '5.9':  '部分可自动化（check_network 提取设备型号，异构部署结论需比对台账）',
    '5.12': '部分可自动化（check_network 采集 IKE SA 识别网络层加密，加密层次组合需按方案人工核对）',
    '5.14': '部分可自动化（check_network 核对防火墙会话表比对跨域通道，隔离方式与等级对应需人工核查）',
    '5.17': '部分可自动化（check_network 采集 IKE SA 识别网络层加密，链路层/信源层加密需按方案人工核对）',
    '5.22': '部分可自动化（check_network 提取存储网/业务网 VLAN 划分信号，结论需比对台账）',
}
# 需人工（脚本未执行该项检测，仅输出方法提示/属性使然）
HAND_MANUAL = {
    '3.4':  '需人工（载体物理销毁，脚本不适用；核查销毁档案与影像记录）',
    '3.12': '需人工（采集范围与业务需求比对，脚本未执行检测）',
    '3.14': '需人工（大数据平台统一管控，需平台界面人工核查）',
    '4.6':  '需人工（Web 应用防护与应用层输入校验需人工核查；数据库连接最小权限脚本未执行该项检测）',
    '4.9':  '需人工（SQL注入/XSS防护需 WAF/代码审计，脚本仅输出方法提示）',
    '4.10': '需人工（执行代码校验需 nikto 扫描或人工，脚本仅输出方法提示）',
    '4.13': '需人工（管理终端专设专用需现场核对台账，脚本仅输出方法提示）',
    '4.15': '需人工（签名验证/密级标识功能需人工核查，脚本仅输出方法提示）',
    '4.20': '需人工（自主协议/接口认定需人工核查，脚本仅输出方法提示）',
    '4.21': '需人工（多因素与数字证书认证需人工核查，脚本仅输出方法提示）',
    '4.26': '需人工（代码级漏洞挖掘需渗透测试/代码审计，脚本仅输出方法提示）',
    '4.27': '需人工（输入格式长度校验有效性需人工测试，脚本仅输出方法提示）',
    '4.28': '需人工（四类攻击防御能力需 WAF/渗透测试验证，脚本仅输出方法提示）',
    '4.29': '需人工（统一权限管理平台需人工核查 IAM/SSO，脚本仅输出方法提示）',
}

def caps_of(code):
    """该行勾选矩阵对应的 [(对象类型, 脚本能力)]，按 TYPE_ORDER 排序。"""
    appl = applic_map.get(code, [])
    return [(t, target_cap(code, t)) for t in TYPE_ORDER if t in appl]

def gen_auto_text(code):
    autos = [t for t, c in caps_of(code) if c == 'auto']
    if not autos:
        return None
    return '可自动化（用 ' + '、'.join(SCRIPT_NAME[t] for t in autos) + ' 自动核查）'

def gen_partial_text(code):
    caps = caps_of(code)
    autos = [t for t, c in caps if c == 'auto']
    checks = [t for t, c in caps if c == 'check']
    rems = [t for t, c in caps if c not in ('auto', 'check')]
    parts = []
    if autos:
        parts.append('用 ' + '、'.join(SCRIPT_NAME[t] for t in autos) + ' 自动核查')
    if checks:
        parts.append('用 ' + '、'.join(SCRIPT_NAME[t] for t in checks) + ' 采集检测结果、结论需人工确认')
    if rems:
        parts.append('、'.join(SCRIPT_NAME[t] for t in rems) + ' 未执行该项检测，需人工核查')
    return '部分可自动化（' + '；'.join(parts) + '）'

def gen_manual_text(code):
    rems = [t for t, c in caps_of(code) if c not in ('auto', 'check')]
    if rems:
        return '需人工（' + '、'.join(SCRIPT_NAME[t] for t in rems) + ' 未执行检测，需按指导书方法人工核查）'
    return '需人工（脚本未执行检测，需按指导书方法人工核查）'

def compose(code):
    ch = int(code.split('.')[0])
    if code in HAND_AUTO:
        return HAND_AUTO[code]
    if code in HAND_PARTIAL:
        return HAND_PARTIAL[code]
    if code in HAND_MANUAL:
        return HAND_MANUAL[code]
    if ch == 5:
        # 第5章主体为网络设备（个别行含麒麟：如 5.19 终端逻辑隔离）
        caps = caps_of(code)
        net_auto = cap.get(code, {}).get('kylin/check_network', 'none') == 'auto'
        other_ok = all(c == 'auto' for t, c in caps if t != '网络设备')
        if net_auto and other_ok:
            return '可自动化（用 check_network.sh / check_network.ps1 采集设备回显后自动判定）'
        if any(c in ('auto', 'check') for _, c in caps):
            return gen_partial_text(code)
        return '需人工（台账/平台核查）'
    if ch in (6, 7, 8, 9):
        return '需人工（现场/文档/台账类核查）'
    if ch == 10:
        return None                     # 10.1 由 HAND_AUTO 提供
    caps = caps_of(code)
    if not caps:
        return None
    if all(c == 'auto' for _, c in caps):
        return gen_auto_text(code)
    if any(c in ('auto', 'check') for _, c in caps):
        return gen_partial_text(code)
    return gen_manual_text(code)

# ---------- 4. 分类校验 ----------
def classify(code):
    txt = compose(code)
    if txt is None:
        return None
    return ('可自动化' if txt.startswith('可自动化')
            else ('部分可自动化' if txt.startswith('部分可自动化') else '需人工'))

texts, errors = [], []
dist = Counter()
for r in range(4, 140):
    code = rowmap[r]
    txt = compose(code)
    caps = dict(caps_of(code))
    if txt is None:
        errors.append(f'R{r} {code}: 无文案，对象能力={caps}')
        continue
    n_auto = sum(1 for c in caps.values() if c == 'auto')
    n_mixed = sum(1 for c in caps.values() if c in ('auto', 'check'))
    if txt.startswith('可自动化') and code != '10.1' and (not caps or n_auto != len(caps)):
        errors.append(f'R{r} {code}: 标可自动化但能力分布={caps}')
    if txt.startswith('部分可自动化') and n_mixed == 0:
        errors.append(f'R{r} {code}: 标部分可自动化但无 auto/check 对象 能力={caps}')
    if txt.startswith('需人工') and n_mixed > 0:
        errors.append(f'R{r} {code}: 标需人工但存在 auto/check 对象 能力={caps}')
    # 「部分可自动化」的行：未自动化的适用对象必须能被使用方看出来（点名脚本或所属类别）
    if txt.startswith('部分可自动化'):
        for t, c in caps.items():
            if c == 'auto':
                continue
            if not any(tok in txt for tok in FAMILY[t]):
                errors.append(f'R{r} {code}: 适用对象「{t}」能力={c} 未在文案中点名 -> {txt}')
    # 文案断言某对象"未执行"检测，但实测该对象脚本确已执行（auto/check）-> 口径自相矛盾
    if '未执行' in txt:
        for seg in re.split(r'[（）；;，,]', txt):
            if '未执行' not in seg:
                continue
            for t, c in caps.items():
                if c in ('auto', 'check') and any(tok in seg for tok in FAMILY[t]):
                    errors.append(f'R{r} {code}: 文案称「{t}」未执行，但实测能力={c} -> {seg.strip()}')
    texts.append((r, txt))
    dist[classify(code)] += 1

if errors:
    print('!! 校验失败：')
    for e in errors:
        print('  ', e)
    sys.exit(1)

EXPECT = {'可自动化': 43, '部分可自动化': 51, '需人工': 42}
print(f'=== 干跑：{len(texts)} 条，分布 {dict(dist)}，预期 {EXPECT} ===')
if '--show' in sys.argv:
    for r, txt in texts:
        caps = dict(caps_of(rowmap[r]))
        print(f'  R{r:3d} {rowmap[r]:5s} {txt}')
        print(f'        能力 {caps}')
if not APPLY:
    print('（干跑模式，未写盘。加 --apply 应用）')
    sys.exit(0)
if dict(dist) != EXPECT:
    print('!! 分布与预期不符，中止（如口径已确认变更，请同步 EXPECT）')
    sys.exit(1)

# ---------- 5. 写入 ----------
for r, txt in texts:
    ws.cell(row=r, column=14).value = txt
ws.column_dimensions['N'].width = 55
for r, txt in texts:
    c = ws.cell(row=r, column=14)
    c.alignment = Alignment(wrap_text=True, vertical='center')
    # N 列宽 55 ≈ 每行 26 个汉字，行高按行数取足，避免固定行高裁掉标注文字
    lines = math.ceil(len(txt) / 26)
    need = lines * 15 + 6
    h = ws.row_dimensions[r].height
    if h is not None and h < need:
        ws.row_dimensions[r].height = need

wb.save(XLSX)
print(f'已写入 {XLSX}（脚本能否核查口径，{len(texts)} 行标注）')
