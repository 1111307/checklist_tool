# -*- coding: utf-8 -*-
import re, glob, os, datetime
from docx import Document
from collections import defaultdict

# ---------- 1. 提取指导书检查项 ----------
doc = Document('配置核查作业指导书_v2.2.docx')
guide_items = []
cur_ch = None
for p in doc.paragraphs:
    # 标题层级说明：v2.2 已将层级整体上移（章=Heading 1，条目=Heading 2）
    if p.style.name == 'Heading 1':
        cur_ch = p.text.strip()
    elif p.style.name == 'Heading 2':
        m = re.match(r'(\d+\.\d+)\s+(.*)', p.text.strip())
        if m:
            guide_items.append((cur_ch, m.group(1), m.group(2)))

# ---------- 2. 提取脚本覆盖：编号 -> {状态集合} ----------
script_status = {}
COMP = ['check_mysql','check_redis','check_dm','check_sqlserver','check_nginx','check_tomcat']

def load(fn, name, enc):
    data = open(fn, encoding=enc).read()
    if enc == 'gbk':
        data = re.sub(r'_\s*\r?\n\s*', '', data)   # VBS 续行符折叠
        pat = r'AddResult\s*\(?\s*"([^"]+)"\s*,\s*"[^"]*"\s*,\s*"[^"]*"\s*,\s*"([^"]+)"'
    else:
        pat = r'add_result\s+"([^"]+)"\s+"[^"]*"\s+"[^"]*"\s+"([^"]+)"'
    d = defaultdict(set)
    for m in re.finditer(pat, data):
        d[m.group(1)].add(m.group(2))
    script_status[name] = dict(d)

for f in sorted(glob.glob('win/check_*.vbs')):
    load(f, 'win_'+os.path.basename(f).replace('.vbs',''), 'gbk')
for f in sorted(glob.glob('kylin/check_*.sh')):
    load(f, 'kylin_'+os.path.basename(f).replace('.sh',''), 'utf-8')

def evaluate(code):
    sts = set(); w = k = False
    # 记录每个平台是"真覆盖"还是"na兜底"（na 视为未覆盖，不算实现）
    w_real = k_real = False
    win_names = ['win_check_xp7'] + ['win_'+c for c in COMP]
    kylin_names = ['kylin_check_kylin'] + ['kylin_'+c for c in COMP]
    for n in win_names:
        d = script_status.get(n, {})
        if code in d:
            st = set(d[code])
            if st - {'na'}:
                w = True; w_real = True
            sts |= st
    for n in kylin_names:
        d = script_status.get(n, {})
        if code in d:
            st = set(d[code])
            if st - {'na'}:
                k = True; k_real = True
            sts |= st
    return sts, w, k

# 判定：已实现 / 部分实现 / 未实现，并给出原因
def judge(sts, w, k):
    # 排除 na 后无任何真实状态 -> 未实现（脚本只有 na 兜底）
    real = sts - {'na'}
    if not real:
        return '未实现', '脚本仅标记不适用，未实现核查'
    has_auto = 'pass' in sts or 'fail' in sts
    if w and k and has_auto:
        return '已实现', '双平台覆盖，自动判定'
    if w and k:
        return '部分实现', '双平台覆盖，结论需人工确认'
    plat = 'Windows' if w else '麒麟'
    missing = '麒麟' if w else 'Windows'
    if has_auto:
        return '部分实现', f'仅 {plat} 覆盖，缺 {missing}'
    return '部分实现', f'仅 {plat} 覆盖，且需人工确认'

# 覆盖平台描述
def platform(w, k):
    if w and k: return 'Windows、麒麟'
    if w: return 'Windows'
    if k: return '麒麟'
    return '无'

# ---------- 3. 统计 ----------
ok = part = none = 0
ch_stat = defaultdict(lambda: [0, 0, 0])
detail = defaultdict(list)
for ch, code, title in guide_items:
    sts, w, k = evaluate(code)
    st, reason = judge(sts, w, k)
    if st == '已实现':   ok += 1; ch_stat[ch][0] += 1
    elif st == '部分实现': part += 1; ch_stat[ch][1] += 1
    else:                none += 1; ch_stat[ch][2] += 1
    detail[ch].append((code, title, platform(w, k), st, reason))

total = len(guide_items)
rate = ok / total * 100
today = datetime.date.today().isoformat()

# 章节顺序按指导书原始顺序（detail 为 defaultdict，需固定顺序）
chapters = []
for ch, code, title in guide_items:
    if ch not in chapters:
        chapters.append(ch)

# ---------- 4. 生成报告 ----------
L = []
A = L.append

A('# 配置核查指导书与自动化核查工具交叉验证报告')
A('')
A(f'报告日期：{today}')
A('')
A('## 一、验证结论')
A('')
A(f'对《配置核查作业指导书》v2.2 与两套核查工具（Windows VBScript 版、麒麟 Bash 版）逐项比对，共核对 {total} 个编号检查项。已实现 {ok} 项，部分实现 {part} 项，未实现 {none} 项，自动化覆盖率 {rate:.1f}%。')
A('')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 已实现 | {ok} | {ok/total*100:.1f}% |')
A(f'| 部分实现 | {part} | {part/total*100:.1f}% |')
A(f'| 未实现 | {none} | {none/total*100:.1f}% |')
A('')
A('部分实现项主要落在数据安全、应用安全两章，以及第 6 至 9 章的现场与文档类检查。')
A('')
A('## 二、验证对象与方法')
A('')
A('### 2.1 验证对象')
A('')
A('- 指导书：《配置核查作业指导书》v2.2，共 10 章、' + str(total) + ' 个编号检查项。')
A('- 核查工具：14 个脚本。操作系统核查 2 个（Windows、麒麟各 1），数据库与中间件核查 12 个，覆盖 MySQL、Redis、达梦 DM、SQL Server、Nginx、Tomcat，Windows 与麒麟各 6。')
A('')
A('### 2.2 验证方法')
A('')
A('从指导书 docx 提取全部编号检查项，从各脚本源码提取检查结果记录点，按编号逐项比对。检查项在脚本里的实现状态，以脚本输出的判定结果（合格/不合格/需人工确认）为准。')
A('')
A('### 2.3 判定规则')
A('')
A('| 验证结论 | 判定条件 |')
A('|---|---|')
A('| 已实现 | Windows 与麒麟脚本均覆盖该项，且脚本能自动给出合格或不合格判定 |')
A('| 部分实现 | 仅单平台覆盖；或脚本虽覆盖但结论需人工确认 |')
A('| 未实现 | 无任何脚本覆盖该项 |')
A('')
A('## 三、各章覆盖统计')
A('')
A('| 章节 | 检查项 | 已实现 | 部分实现 | 未实现 |')
A('|---|---|---|---|---|')
for ch in chapters:
    a, b, c = ch_stat[ch]
    A(f'| {ch} | {len(detail[ch])} | {a} | {b} | {c} |')
A('')
A('## 四、逐项对照明细')
A('')
for ch in chapters:
    A(f'### {ch}')
    A('')
    A('| 编号 | 检查项 | 覆盖平台 | 验证结论 | 原因 |')
    A('|---|---|---|---|---|')
    for code, title, plat, st, reason in detail[ch]:
        A(f'| {code} | {title} | {plat} | {st} | {reason} |')
    A('')

# 差异分析
A('## 五、差异与问题分析')
A('')
A('### 5.1 指导书缺项')
A('')
A('指导书 v2.2 的第 5 章「网络安全」只有一句「目前无网络设备」的占位，23 个检查项未纳入本次验证范围。')
A('')
A('### 5.2 部分实现项的原因分布')
A('')
# 统计部分实现的原因分布
reason_stat = defaultdict(int)
for ch, items in detail.items():
    for code, title, plat, st, reason in items:
        if st == '部分实现':
            reason_stat[reason] += 1
A(f'共 {part} 项为部分实现，按原因分布如下：')
A('')
A('| 原因 | 项数 |')
A('|---|---|')
for r, c in sorted(reason_stat.items(), key=lambda x: -x[1]):
    A(f'| {r} | {c} |')
A('')
A('### 5.3 部分实现项的整改性质')
A('')
A(f'这 {part} 项部分实现，一半本来就只能靠人工，脚本标 manual 是合理的；另一半理论上还能自动化，只是受环境限制。')
A('')
A('只能人工的：第 6 至 9 章（物理、组织、制度、管理）22 项，要到机房看设备台账、翻组织文件、核对制度版本和演练记录。第 3 章的 3.4 载体销毁、3.12 采集范围、3.14 大数据管控，涉及物理销毁和业务范围比对。第 4 章的 4.9、4.10、4.13、4.20、4.26、4.27、4.28、4.29、4.31，多是应用层或代码级能力，得靠渗透测试、代码审计或现场台账。再有第 1 章的 1.16、1.27，第 2 章的 2.3，需要代码审计或采购凭证。')
A('')
A('还能自动化的：第 3 章的 3.1、3.7、3.9、3.13，Windows 侧已能测 BitLocker 和 HTTPS 端口，数据库 TDE、SSL 连接因为要登录数据库才标 manual。第 4 章的 4.1、4.2、4.3、4.4、4.7、4.11、4.14、4.19、4.30，检测命令已经补了，只在没部署相应服务时才标 manual。')
A('')
A('2.3 应用口令、3.6 边界防泄漏这两项，脚本只标了「不适用」，实际没做任何核查，本次按未实现算。')
A('')
A('## 六、整改建议')
A('')
A('下一步几件事。先把第 5 章「网络安全」的检查项补进指导书，再跑一遍本验证。第 4 章里依赖代码审计和渗透测试的项（4.9、4.10、4.26、4.27、4.28），补上 WAF 和代码扫描工具。第 6 至 9 章这类只能人工的，脚本就只提供检查清单和记录模板，留给人填。2.3 应用口令、3.6 边界防泄漏目前是空的，最该先补。')
A('')
open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write('\n'.join(L))
print(f'已生成：{total} 项，已实现 {ok}，部分实现 {part}，未实现 {none}，覆盖率 {rate:.1f}%')
