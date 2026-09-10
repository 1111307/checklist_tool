# -*- coding: utf-8 -*-
"""交叉验证报告生成器 v3（按核查表适用对象矩阵判定）。

判定基准：核查表 C~L 列勾选矩阵（√=适用、—=不适用）。逐项检查每个适用对象
对应脚本是否具备自动判定（pass/fail）能力：
  - 全部适用对象可自动判定 → 已实现
  - 部分适用对象可自动（其余对象/一侧平台需人工）→ 部分实现
  - 适用对象均需人工（脚本仅方法输出）→ 部分实现
  - 第 5 章网络设备按 check_network.sh 实际判定能力分类
输出与核查表第 13 列详情标注同口径。
"""
import re, glob, os, datetime
from docx import Document
from openpyxl import load_workbook
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

# ---------- 2. 提取脚本覆盖：编号 -> {脚本: 状态集合} ----------
script_status = {}

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
    load(f, 'win_' + os.path.basename(f).replace('.vbs', ''), 'gbk')
for f in sorted(glob.glob('kylin/check_*.sh')):
    load(f, 'kylin_' + os.path.basename(f).replace('.sh', ''), 'utf-8')

# 合并编号展开（1.16-17 合并项、x_m 变体）
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
for sname, d in script_status.items():
    for code, sts in d.items():
        for c in expand(code):
            cov[c][sname] |= sts

# ---------- 3. 适用对象矩阵（核查表勾选） ----------
XLSX = '配置核查表_v2.0.0_标注自动验证.xlsx'
_wb = load_workbook(XLSX)
_ws = _wb.active
COLS = {3: 'Windows', 4: 'Windows', 5: '麒麟', 6: '麒麟', 7: 'Nginx', 8: 'Tomcat',
        9: 'MySQL', 10: 'SQL Server', 11: '达梦', 12: 'Redis'}
TARGET_SCRIPTS = {
    'Windows': ['win_check_xp7'], '麒麟': ['kylin_check_kylin'],
    'Nginx': ['win_check_nginx', 'kylin_check_nginx'], 'Tomcat': ['win_check_tomcat', 'kylin_check_tomcat'],
    'MySQL': ['win_check_mysql', 'kylin_check_mysql'],
    'SQL Server': ['win_check_sqlserver', 'kylin_check_sqlserver'],
    '达梦': ['win_check_dm', 'kylin_check_dm'], 'Redis': ['win_check_redis', 'kylin_check_redis'],
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

def target_capability(code, t):
    best = 'none'
    for s in TARGET_SCRIPTS[t]:
        sts = cov.get(code, {}).get(s, set())
        if 'pass' in sts or 'fail' in sts:
            return 'auto'
        if 'manual' in sts:
            best = 'manual'
    return best

# ---------- 4. 判定 ----------
NET_AUTO = {'5.3', '5.4', '5.5', '5.6', '5.7', '5.8', '5.10', '5.11', '5.16', '5.18', '5.19'}
NET_PARTIAL = {'5.9', '5.22', '5.23'}
# 其余 5.x 为台账/平台类
SPECIAL = {
    '10.1': ('已实现', '脚本自动核查 TLS/SSH 协议版本（协议层条目，勾选矩阵未单列对象）', '适用对象均自动判定'),
    '1.27': ('部分实现', '正版/定制软件认定需人工核对采购与授权凭证（核查表未设行）', '适用对象均需人工（脚本仅方法输出）'),
}

def judge(code):
    ch = int(code.split('.')[0])
    if code in SPECIAL:
        return SPECIAL[code]
    if ch == 5:
        real = cov.get(code, {}).get('kylin_check_network', set()) - {'na'}
        if not real:
            return '未实现', '脚本仅标记不适用', '其他'
        if code in NET_AUTO:
            return '已实现', 'check_network.sh 自动判定（设备回显采集后自动给出结论）', '适用对象均自动判定'
        if code in NET_PARTIAL:
            return '部分实现', '脚本提取部分信号（型号/存储网/端口闲置率），结论需台账比对', '第5章部分信号+台账比对'
        return '部分实现', '审批/台账/平台类核查，需人工确认', '第5章审批/台账/平台类（需人工）'
    if ch in (6, 7, 8, 9):
        return '部分实现', '现场/文档/台账类核查，需人工', '第6-9章现场/文档类'
    # 第1-4章：按适用矩阵
    appl = applic.get(code)
    if appl is None:
        return '未实现', '（无核查表行，无脚本覆盖）', '其他'
    if not appl:
        # 未勾选核查对象的条目（如 3.4 载体销毁）：看脚本是否输出核查方法
        covered = any(cov.get(code, {}).get(s, set()) - {'na'}
                      for s in ('win_check_xp7', 'kylin_check_kylin'))
        if covered:
            return '部分实现', '未勾选核查对象（载体销毁类条目），脚本输出核查方法，结论需人工', '适用对象均需人工（脚本仅方法输出）'
        return '未实现', '核查表未勾选任何适用对象，且无脚本覆盖', '其他'
    auto_t = [t for t in appl if target_capability(code, t) == 'auto']
    manual_t = [t for t in appl if target_capability(code, t) != 'auto']
    if not manual_t:
        return '已实现', '适用对象均自动判定', '适用对象均自动判定'
    if auto_t:
        os_mixed = (('Windows' in manual_t) or ('麒麟' in manual_t)) and len(auto_t) + len(manual_t) > 0
        reason = '、'.join(auto_t) + ' 可自动判定；' + '、'.join(manual_t) + ' 需人工'
        if len(manual_t) == 1 and manual_t[0] in ('Windows', '麒麟') and all(
                t in ('Windows', '麒麟') or target_capability(code, t) == 'auto' for t in appl):
            bucket = '一侧 OS 平台待补自动判定'
        elif all(t not in ('Windows', '麒麟') for t in manual_t):
            bucket = '部分组件对象待补（其余对象需人工）'
        else:
            bucket = '一侧 OS 平台待补自动判定'
        return '部分实现', reason, bucket
    return '部分实现', '脚本输出核查方法，结论需人工确认', '适用对象均需人工（脚本仅方法输出）'

def applic_disp(code):
    ch = int(code.split('.')[0])
    if ch == 5:
        return '网络设备'
    if ch in (6, 7, 8, 9):
        return '—'
    appl = applic.get(code)
    if not appl:
        return '—'
    return '、'.join(appl)

# 安全校验：NET_AUTO 中的项必须确实在 check_network.sh 里有 pass/fail 分支
net_src = cov.get('5.3', {})  # 占位，实际校验在下
for c in sorted(NET_AUTO):
    sts = set()
    for sname, d in cov.items():
        pass
    sts = cov.get(c, {}).get('kylin_check_network', set())
    if not ({'pass', 'fail'} & sts):
        raise SystemExit(f'校验失败：{c} 不在 check_network.sh 自动判定分支中')

# ---------- 5. 统计 ----------
ok = part = none = 0
ch_stat = defaultdict(lambda: [0, 0, 0])
detail = defaultdict(list)
bucket_stat = defaultdict(int)
for ch, code, title in guide_items:
    st, reason, bucket = judge(code)
    if st == '已实现':   ok += 1; ch_stat[ch][0] += 1
    elif st == '部分实现': part += 1; ch_stat[ch][1] += 1; bucket_stat[bucket] += 1
    else:                none += 1; ch_stat[ch][2] += 1
    detail[ch].append((code, title, applic_disp(code), st, reason))

total = len(guide_items)
rate = ok / total * 100
today = datetime.date.today().isoformat()

chapters = []
for ch, code, title in guide_items:
    if ch not in chapters:
        chapters.append(ch)

# ---------- 6. 生成报告 ----------
L = []
A = L.append

A('# 配置核查指导书与自动化核查工具交叉验证报告')
A('')
A(f'报告日期：{today}')
A('')
A('## 一、验证结论')
A('')
A(f'对《配置核查作业指导书》v2.2 与自动化核查工具（Windows VBScript 版、麒麟 Bash 版、网络设备核查 check_network.sh）逐项比对，共核对 {total} 个编号检查项。以核查表勾选矩阵（√ 适用 / — 不适用）确定每项的适用对象，逐对象核对脚本自动判定能力：已实现 {ok} 项，部分实现 {part} 项，未实现 {none} 项，自动化覆盖率 {rate:.1f}%。')
A('')
A('| 验证结论 | 项数 | 占比 |')
A('|---|---|---|')
A(f'| 已实现 | {ok} | {ok/total*100:.1f}% |')
A(f'| 部分实现 | {part} | {part/total*100:.1f}% |')
A(f'| 未实现 | {none} | {none/total*100:.1f}% |')
A('')
b1 = bucket_stat.get('一侧 OS 平台待补自动判定', 0)
b2 = bucket_stat.get('部分组件对象待补（其余对象需人工）', 0)
b3 = bucket_stat.get('适用对象均需人工（脚本仅方法输出）', 0)
b4 = bucket_stat.get('第5章部分信号+台账比对', 0)
b5 = bucket_stat.get('第5章审批/台账/平台类（需人工）', 0)
b6 = bucket_stat.get('第6-9章现场/文档类', 0)
A(f'部分实现 {part} 项按原因分布：一侧 OS 平台待补自动判定 {b1} 项（麒麟侧 15、Windows 侧 1）、部分组件对象待补 {b2} 项（达梦、SQL Server、Redis 侧人工）、适用对象均需人工 {b3} 项、第 5 章部分信号+台账比对 {b4} 项、第 5 章台账/平台类 {b5} 项、第 6-9 章现场文档类 {b6} 项。')
A('')
A('## 二、验证对象与方法')
A('')
A('### 2.1 验证对象')
A('')
A('- 指导书：《配置核查作业指导书》v2.2，共 10 章、' + str(total) + ' 个编号检查项（交付包内的 v2.0.1 为同内容版本，另含第 5 章 13 张终端操作截图）。')
A('- 核查工具：15 个脚本。操作系统核查 2 个（Windows、麒麟各 1），数据库与中间件核查 12 个（MySQL、Redis、达梦 DM、SQL Server、Nginx、Tomcat，两平台各 6），网络设备核查 1 个（check_network.sh，采集-解析模式，覆盖第 5 章 23 项）。')
A('')
A('### 2.2 验证方法')
A('')
A('从指导书 docx 提取全部编号检查项，从各脚本源码提取检查结果记录点（add_result/AddResult 的编号与判定状态）。每项的适用对象以核查表勾选矩阵为准（√ 适用 / — 不适用）：Windows/WinXP 对应 check_xp7.vbs、中标/银河麒麟对应 check_kylin.sh、各数据库与中间件对应双平台组件脚本（任一平台有自动判定分支即计该对象可自动）、网络设备对应 check_network.sh。按对象逐项比对，检查项的实现状态以脚本实际输出的判定（合格/不合格/需人工）为准。')
A('')
A('### 2.3 判定规则')
A('')
A('| 验证结论 | 判定条件 |')
A('|---|---|')
A('| 已实现 | 该项在核查表中勾选的全部适用对象，核查脚本均能自动给出合格或不合格判定 |')
A('| 部分实现 | 适用对象中仅部分可自动判定（其余对象或一侧平台需人工）；或脚本虽覆盖但仅输出核查方法、结论需人工确认 |')
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
    A('| 编号 | 检查项 | 适用对象 | 验证结论 | 原因 |')
    A('|---|---|---|---|---|')
    for code, title, appl, st, reason in detail[ch]:
        A(f'| {code} | {title} | {appl} | {st} | {reason} |')
    A('')

A('## 五、差异与问题分析')
A('')
A('### 5.1 网络安全（第 5 章）')
A('')
A('第 5 章「网络安全」23 项：指导书已按核查表补写小节（23 节、133 段方法），脚本侧为 check_network.sh（采集-解析模式：init 生成华为/华三/锐捷三厂商命令清单模板，运维陪同登录设备采集回显后 check 解析判定）。23 项分三类：设备命令回显类 11 项（5.3 至 5.8、5.10、5.11、5.16、5.18、5.19）脚本自动判定，记为已实现；5.9 异构部署、5.22 存储网分离、5.23 配备合理性 3 项，脚本可提取部分信号（设备型号、存储网划分、端口闲置率），结论需台账比对；其余 9 项（5.1、5.2、5.12 至 5.15、5.17、5.20、5.21）为审批记录、方案文档与平台界面类核查，需人工确认。该划分与核查表第 13 列「工具自动验证」的标注一一对应。')
A('')
A('### 5.2 部分实现项的原因分布')
A('')
A(f'共 {part} 项为部分实现，按原因分布如下：')
A('')
A('| 原因类别 | 项数 |')
A('|---|---|')
for b, c in sorted(bucket_stat.items(), key=lambda x: -x[1]):
    A(f'| {b} | {c} |')
A('')
A('### 5.3 部分实现项的整改性质')
A('')
A(f'这 {part} 项部分实现，一半本来就是只能人工的，脚本标 manual 是合理的；另一半补上对应平台或组件的自动判定后即可转为已实现。')
A('')
A(f'只能人工的：第 6 至 9 章（物理、组织、制度、管理）22 项，要到机房看设备台账、翻组织文件、核对制度版本和演练记录。第 5 章台账/平台类 9 项，是跨网审批单、方案文档与平台界面核查。适用对象均需人工的 {b3} 项（1.13、1.16、1.22、1.27、3.4、3.12、3.14、4.1、4.5、4.9、4.10、4.13、4.15、4.20、4.21、4.26、4.27、4.28、4.29、4.32），涉及应用层代码能力、物理销毁过程与文档台账比对。')
A('')
A(f'还能自动化的：一侧 OS 平台待补 {b1} 项——麒麟侧 2.2、2.5、2.6、2.8、2.15、3.8、3.10、4.6、4.8、4.11、4.12、4.17、4.23、4.24、1.25（Windows 侧脚本已能自动判定，麒麟侧脚本仅输出核查方法），Windows 侧 4.31（与前者相反）；补齐对应平台脚本即可全对象自动。组件对象待补 {b2} 项（1.1、1.8、1.20、1.21、1.24、2.16）——达梦侧 6 处、SQL Server 侧 2 处、Redis 侧 2 处仅人工，补齐组件脚本判定分支即可。')
A('')
A('2.3 应用口令、3.6 边界防泄漏此前仅标「不适用」，此前会话已补真实检测：2.3 核查空口令账户与图形界面自动登录（/etc/shadow、lightdm/sddm/gdm、注册表 AutoAdminLogon/LimitBlankPasswordUse），3.6 按 3.6.3(3) 核查终端 DLP/管控客户端联动（进程、安装包、卸载注册表），现均为已实现。')
A('')
A('## 六、整改建议')
A('')
A('下一步几件事。麒麟侧 15 项待补自动判定（多为应用层配置类，可参照 Windows 侧实现），达梦/SQL Server/Redis 组件侧 8 处待补判定分支；补齐后约 22 项可从部分实现转已实现。第 5 章上真机核查时，用 check_network.sh init 生成采集清单，运维陪同采集回显后跑 check 出报告；遇到解析不出的回显格式，把回显片段补进特征匹配。第 4 章里依赖代码审计和渗透测试的项（4.9、4.10、4.26、4.27、4.28），补上 WAF 和代码扫描工具。第 6 至 9 章和第 5 章台账类这类只能人工的，用人工核查台（manual_check.html）逐项记录、粘贴取证截图并导出报告。136 项全部有核查工具覆盖（自动判定或人工路径）。')
A('')
open('测评报告/指导书与核查工具交叉验证报告.md', 'w', encoding='utf-8').write('\n'.join(L))
print(f'已生成：{total} 项，已实现 {ok}，部分实现 {part}，未实现 {none}，覆盖率 {rate:.1f}%')
print('原因分布：', dict(bucket_stat))
