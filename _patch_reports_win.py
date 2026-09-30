# -*- coding: utf-8 -*-
"""评审整改第 4/5/8/9 条（报告结构）——Windows 侧的统一改法（与 _patch_reports_kylin.py 同口径）。

改什么（HTML / XLS / XLSX / CSV / JSON 五个输出面 + Add-Result 签名）：
  4. 加一栏「验证过程/方法」：放在 结果 与 详情 之间（评审第 9 条：状态与详情中间要有步骤）。
     数据源 = _method_map.json（指导书 v2.0.0 逐条提取）。
       - PS1 脚本（6 个组件 + check_network）：dot-source 新文件 win/lib_method.ps1（哈希表查询）。
       - check_xp7.vbs（GBK）：把 136 条内嵌成 MethodOf / GuideTitleOf 两个 Select Case（单文件零依赖）。
  5. 「建议」栏拆两栏：不合规(fail)→「修改建议」，其余(含合规)→「安全要求」。
  8. 核查项名对齐指导书：优先取指导书 H2 标题；差异时脚本项名以〔脚本项名：…〕并入详情前缀留痕。

涉及文件：
  win/lib_method.ps1（新建）  win/check_{mysql,redis,dm,sqlserver,nginx,tomcat}.ps1
  win/lib_xlsx.ps1            win/check_network.ps1            win/check_xp7.vbs（GBK，Python 读写）
"""
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

METHOD = json.load(open("_method_map.json", encoding="utf-8"))
KEYS = sorted(METHOD, key=lambda s: [int(x) for x in s.split(".")])

report = []


def ok(msg):
    report.append(msg)


# ============================================================
# A. win/lib_method.ps1（新建，UTF-8 BOM 与其余 ps1 一致）
# ============================================================
def ps_quote(s):
    return "'" + s.replace("'", "''") + "'"


lib_method = """# ============================================================
# 验证过程/方法 与 核查项名 查询表（评审整改 2026-09-30 第4/8条）
# 数据源：配置核查作业指导书 v2.0.0 逐条提取（_gen_method_map.py 生成 _method_map.json）
# 用法（核查脚本内 dot-source 后调用）：
#   . "$ScriptDir\\lib_method.ps1"
#   Method-Of "1.4"        # 该条验证过程/方法（指导书口径）
#   Guide-Title-Of "1.4"   # 指导书 H2 标题（核查项名对齐用）
# ============================================================
$script:METHOD_MAP = @{
"""
for k in KEYS:
    lib_method += "    %s = %s\n" % (ps_quote(k), ps_quote(METHOD[k]["method"]))
lib_method += """}
$script:GUIDE_TITLE_MAP = @{
"""
for k in KEYS:
    lib_method += "    %s = %s\n" % (ps_quote(k), ps_quote(METHOD[k]["title"]))
lib_method += """}
function Method-Of([string]$id) {
    if ($script:METHOD_MAP.ContainsKey($id)) { return $script:METHOD_MAP[$id] } else { return '' }
}
function Guide-Title-Of([string]$id) {
    if ($script:GUIDE_TITLE_MAP.ContainsKey($id)) { return $script:GUIDE_TITLE_MAP[$id] } else { return '' }
}
"""
open("win/lib_method.ps1", "w", encoding="utf-8-sig", newline="").write(lib_method.replace("\n", "\r\n"))
ok("win/lib_method.ps1: 新建（method %d 条、title %d 条）" % (len(KEYS), len(KEYS)))

# ============================================================
# B. 6 个组件 PS1 脚本
# ============================================================
PS_FILES = ["win/check_mysql.ps1", "win/check_redis.ps1", "win/check_dm.ps1",
            "win/check_sqlserver.ps1", "win/check_nginx.ps1", "win/check_tomcat.ps1"]

PS_LIB_OLD = 'if (Test-Path "$ScriptDir\\lib_xlsx.ps1") { . "$ScriptDir\\lib_xlsx.ps1" }'
PS_LIB_NEW = PS_LIB_OLD + """
# 评审整改（2026-09-30 第4/8条）：验证过程/方法 + 核查项名对齐指导书（lib_method.ps1，指导书 v2.0.0 逐条提取）
if (Test-Path "$ScriptDir\\lib_method.ps1") { . "$ScriptDir\\lib_method.ps1" }"""

PS_ADDRESULT_OLD = """function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += [PSCustomObject]@{ Id=$id; Cat=$cat; Title=$title; Status=$status; Detail=$detail; Chapter=$chapter; Rec=$rec }
}"""

PS_ADDRESULT_NEW = """function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    # 评审整改（2026-09-30 第4/5/8条）：Method=验证过程/方法（指导书逐条）；
    # 核查项名对齐指导书（不一致时以指导书为准，脚本项名以〔脚本项名：…〕并入详情前缀）
    $mt = ''
    if (Get-Command Method-Of -ErrorAction SilentlyContinue) {
        $mt = Method-Of $id
        $gt = Guide-Title-Of $id
        if ($gt -and $gt -ne $title) { $detail = "〔脚本项名：$title〕$detail"; $title = $gt }
    }
    $script:R += [PSCustomObject]@{ Id=$id; Cat=$cat; Title=$title; Status=$status; Method=$mt; Detail=$detail; Chapter=$chapter; Rec=$rec }
}"""

PS_ROW_OLD = """        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.Chapter)</td><td>$(Html-Esc $r.Id)</td><td>$(Html-Esc $r.Cat)</td><td>$(Html-Esc $r.Title)</td><td style='color:$color;font-weight:bold;'>$scn</td><td>$(Html-Esc $r.Detail)</td><td>$(Html-Esc $r.Rec)</td><td>$(Html-Esc (Guide-Ref $r.Id))</td></tr>")"""

PS_ROW_NEW = """        # 评审整改（2026-09-30 第4/5条）：结果之后插验证过程/方法；建议拆修改建议/安全要求两栏
        $mt = if ($null -ne $r.Method) { $r.Method } else { '' }
        $recFix = if ($r.Status -eq 'fail') { $r.Rec } else { '' }
        $reqCol = if ($r.Status -eq 'fail') { '' } else { $r.Rec }
        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.Chapter)</td><td>$(Html-Esc $r.Id)</td><td>$(Html-Esc $r.Cat)</td><td>$(Html-Esc $r.Title)</td><td style='color:$color;font-weight:bold;'>$scn</td><td>$(Html-Esc $mt)</td><td>$(Html-Esc $r.Detail)</td><td>$(Html-Esc $recFix)</td><td>$(Html-Esc $reqCol)</td><td>$(Html-Esc (Guide-Ref $r.Id))</td></tr>")"""

PS_HEAD_OLD = '<th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th>'
PS_HEAD_NEW = '<th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th>'

for p in PS_FILES:
    src = open(p, encoding="utf-8-sig").read()
    nl = "\r\n" if "\r\n" in src else "\n"
    src = src.replace("\r\n", "\n")
    # 幂等：已打补丁（Add-Result 带 Method）则跳过，避免二次插入挂载行
    if "Method=$mt" in src:
        ok("%s: 已打补丁，跳过" % p)
        continue
    changes = []
    if PS_LIB_OLD in src:
        src = src.replace(PS_LIB_OLD, PS_LIB_NEW, 1)
        changes.append("lib_method挂载")
    else:
        ok("!! %s: lib_xlsx 挂载行未匹配" % p)
    if PS_ADDRESULT_OLD in src:
        src = src.replace(PS_ADDRESULT_OLD, PS_ADDRESULT_NEW, 1)
        changes.append("Add-Result")
    else:
        ok("!! %s: Add-Result 未匹配" % p)
    if PS_ROW_OLD in src:
        src = src.replace(PS_ROW_OLD, PS_ROW_NEW, 1)
        changes.append("行构造")
    else:
        ok("!! %s: 行构造未匹配" % p)
    n = src.count(PS_HEAD_OLD)
    src = src.replace(PS_HEAD_OLD, PS_HEAD_NEW)
    changes.append("表头×%d" % n)
    open(p, "w", encoding="utf-8-sig", newline="").write(src.replace("\n", nl))
    ok("%s: %s" % (p, "、".join(changes)))

# ============================================================
# C. win/lib_xlsx.ps1：8 列 → 10 列
# ============================================================
lib = open("win/lib_xlsx.ps1", encoding="utf-8-sig").read()
nl = "\r\n" if "\r\n" in lib else "\n"
lib = lib.replace("\r\n", "\n")
XLSX_HEAD_NEW = "    $hd = @('章节','编号','类别','核查项','结果','验证过程/方法','详情','修改建议','安全要求','参考指导书')"
XLSX_VALS_OLD = """        $guide = if ($GuideScript) { & $GuideScript $r.Id } else { '' }
        $vals = @($r.Chapter, $r.Id, $r.Cat, $r.Title, (Local-StatusCN $r.Status), $r.Detail, $r.Rec, $guide)"""
XLSX_VALS_NEW = """        $guide = if ($GuideScript) { & $GuideScript $r.Id } else { '' }
        # 评审整改（2026-09-30 第4/5条）：验证过程/方法列 + 修改建议/安全要求按状态分栏
        $mt = if ($null -ne $r.Method) { $r.Method } else { '' }
        $recFix = if ($r.Status -eq 'fail') { $r.Rec } else { '' }
        $reqCol = if ($r.Status -eq 'fail') { '' } else { $r.Rec }
        $vals = @($r.Chapter, $r.Id, $r.Cat, $r.Title, (Local-StatusCN $r.Status), $mt, $r.Detail, $recFix, $reqCol, $guide)"""
XLSX_HEAD_OLD = "    $hd = @('章节','编号','类别','核查项','结果','详情','建议','参考指导书')"
ok_lib = []
if "$r.Method" in lib:
    ok("win/lib_xlsx.ps1: 已打补丁，跳过")
else:
    if XLSX_HEAD_OLD in lib:
        lib = lib.replace(XLSX_HEAD_OLD, XLSX_HEAD_NEW, 1)
        ok_lib.append("表头")
    else:
        ok_lib.append("!!表头未匹配")
    if XLSX_VALS_OLD in lib:
        lib = lib.replace(XLSX_VALS_OLD, XLSX_VALS_NEW, 1)
        ok_lib.append("行值")
    else:
        ok_lib.append("!!行值未匹配")
    n = lib.count("for ($c = 0; $c -lt 8; $c++)")
    lib = lib.replace("for ($c = 0; $c -lt 8; $c++)", "for ($c = 0; $c -lt 10; $c++)")
    ok_lib.append("列循环×%d" % n)
    open("win/lib_xlsx.ps1", "w", encoding="utf-8-sig", newline="").write(lib.replace("\n", nl))
    ok("win/lib_xlsx.ps1: %s" % "、".join(ok_lib))

# ============================================================
# D. win/check_network.ps1（哈希表行结构 + JS 表格模板）
# ============================================================
net = open("win/check_network.ps1", encoding="utf-8-sig").read()
nl = "\r\n" if "\r\n" in net else "\n"
net = net.replace("\r\n", "\n")
NET_ADDRESULT_OLD = """function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += @{ id = $id; cat = $cat; title = $title; status = $status; detail = $detail; chapter = $chapter; rec = $rec; guide = "《配置核查作业指导书v2.0.0》第5章 网络安全 $id" }
}"""
NET_ADDRESULT_NEW = """# 评审整改（2026-09-30 第4/8条）：method=验证过程/方法（lib_method.ps1，第5章为采集-解析口径）
if (Test-Path "$ScriptDir\\lib_method.ps1") { . "$ScriptDir\\lib_method.ps1" }
function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    # method=验证过程/方法；核查项名对齐指导书（不一致时以指导书为准，脚本项名以〔脚本项名：…〕并入详情前缀）
    $mt = ''
    if (Get-Command Method-Of -ErrorAction SilentlyContinue) {
        $mt = Method-Of $id
        $gt = Guide-Title-Of $id
        if ($gt -and $gt -ne $title) { $detail = "〔脚本项名：$title〕$detail"; $title = $gt }
    }
    $script:R += @{ id = $id; cat = $cat; title = $title; status = $status; method = $mt; detail = $detail; chapter = $chapter; rec = $rec; guide = "《配置核查作业指导书v2.0.0》第5章 网络安全 $id" }
}"""
NET_JSON_OLD = """        ('{"ch":"' + (Json-Esc $_.chapter) + '","id":"' + (Json-Esc $_.id) + '","cat":"' + (Json-Esc $_.cat) +
         '","title":"' + (Json-Esc $_.title) + '","status":"' + $_.status + '","detail":"' + (Json-Esc $_.detail) +
         '","rec":"' + (Json-Esc $_.rec) + '","guide":"' + (Json-Esc $_.guide) + '"}')"""
NET_JSON_NEW = """        ('{"ch":"' + (Json-Esc $_.chapter) + '","id":"' + (Json-Esc $_.id) + '","cat":"' + (Json-Esc $_.cat) +
         '","title":"' + (Json-Esc $_.title) + '","status":"' + $_.status + '","method":"' + (Json-Esc $_.method) + '","detail":"' + (Json-Esc $_.detail) +
         '","rec":"' + (Json-Esc $_.rec) + '","guide":"' + (Json-Esc $_.guide) + '"}')"""
NET_JS_OLD = """    tr.innerHTML = "<td>"+esc(x.ch)+"</td><td class='id'>"+esc(x.id)+"</td><td class='cat'>"+esc(x.cat)+"</td>"+
      "<td class='title'>"+esc(x.title)+"</td>"+
      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+esc(x.rec)+"</td>"+
      "<td class='guide'>"+esc(x.guide)+"</td>";"""
NET_JS_NEW = """    tr.innerHTML = "<td>"+esc(x.ch)+"</td><td class='id'>"+esc(x.id)+"</td><td class='cat'>"+esc(x.cat)+"</td>"+
      "<td class='title'>"+esc(x.title)+"</td>"+
      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='method'>"+esc(x.method)+"</td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+(x.status=='fail'?esc(x.rec):"")+"</td>"+
      "<td class='req'>"+(x.status=='fail'?"":esc(x.rec))+"</td>"+
      "<td class='guide'>"+esc(x.guide)+"</td>";"""
NET_XLSROW_OLD = """        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.chapter)</td><td>$(Html-Esc $r.id)</td><td>$(Html-Esc $r.cat)</td><td>$(Html-Esc $r.title)</td><td>$scn</td><td>$(Html-Esc $r.detail)</td><td>$(Html-Esc $r.rec)</td><td>$(Html-Esc $r.guide)</td></tr>")"""
NET_XLSROW_NEW = """        # 评审整改（2026-09-30 第4/5条）：结果之后插验证过程/方法；建议拆修改建议/安全要求两栏
        $mt = if ($null -ne $r.method) { $r.method } else { '' }
        $recFix = if ($r.status -eq 'fail') { $r.rec } else { '' }
        $reqCol = if ($r.status -eq 'fail') { '' } else { $r.rec }
        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.chapter)</td><td>$(Html-Esc $r.id)</td><td>$(Html-Esc $r.cat)</td><td>$(Html-Esc $r.title)</td><td>$scn</td><td>$(Html-Esc $mt)</td><td>$(Html-Esc $r.detail)</td><td>$(Html-Esc $recFix)</td><td>$(Html-Esc $reqCol)</td><td>$(Html-Esc $r.guide)</td></tr>")"""
NET_CSS_OLD = "td.detail,td.rec{"
NET_CSS_NEW = "td.detail,td.rec,td.method,td.req{"
if "method = $mt" in net:
    ok("win/check_network.ps1: 已打补丁，跳过")
else:
    ok_net = []
    for old, new, tag, once in [(NET_ADDRESULT_OLD, NET_ADDRESULT_NEW, "Add-Result+lib挂载", True),
                                (NET_JSON_OLD, NET_JSON_NEW, "JSON字段", True),
                                (NET_JS_OLD, NET_JS_NEW, "JS表格", True),
                                (NET_XLSROW_OLD, NET_XLSROW_NEW, "XLS行", True),
                                (PS_HEAD_OLD, PS_HEAD_NEW, "表头", False),
                                (NET_CSS_OLD, NET_CSS_NEW, "CSS", False)]:
        if old in net:
            net = net.replace(old, new, 1) if once else net.replace(old, new)
            n = 1 if once else net.count(new)
            ok_net.append(tag if once else "%s×%d" % (tag, n))
        else:
            ok_net.append("!!%s未匹配" % tag)
    open("win/check_network.ps1", "w", encoding="utf-8-sig", newline="").write(net.replace("\n", nl))
    ok("win/check_network.ps1: %s" % "、".join(ok_net))

# ============================================================
# E. win/check_xp7.vbs（GBK！单文件零依赖：内嵌 Select Case 查询表）
# ============================================================
vbs = open("win/check_xp7.vbs", encoding="gbk").read()


def vbs_quote(s):
    # VBS 字符串不能跨行：方法文本里的换行（多行命令示例）压成单个空格
    import re as _re
    s = _re.sub(r"[\r\n]+", " ", s)
    return '"' + s.replace('"', '""') + '"'


case_method = ["' ============================================================",
               "' 验证过程/方法 查询表（评审整改 2026-09-30 第4条；数据源：指导书 v2.0.0 逐条提取，",
               "' 由 _gen_method_map.py / _patch_reports_win.py 生成，与 win/lib_method.ps1 同源）",
               "' ============================================================",
               "Function MethodOf(id)",
               "    Select Case id"]
for k in KEYS:
    case_method.append('        Case %s: MethodOf = %s' % (vbs_quote(k), vbs_quote(METHOD[k]["method"])))
case_method += ['        Case Else: MethodOf = ""',
                "    End Select",
                "End Function",
                ""]
case_title = ["' ============================================================",
              "' 指导书 H2 标题 查询表（评审整改 2026-09-30 第8条：核查项名对齐指导书）",
              "' ============================================================",
              "Function GuideTitleOf(id)",
              "    Select Case id"]
for k in KEYS:
    case_title.append('        Case %s: GuideTitleOf = %s' % (vbs_quote(k), vbs_quote(METHOD[k]["title"])))
case_title += ['        Case Else: GuideTitleOf = ""',
               "    End Select",
               "End Function",
               ""]
CASE_BLOCK = "\n".join(case_method + case_title)

VBS_ARR_OLD = """Dim rID(400), rCat(400), rTitle(400), rStatus(400)
Dim rDetail(400), rChapter(400), rRec(400)"""
VBS_ARR_NEW = """Dim rID(400), rCat(400), rTitle(400), rStatus(400)
Dim rDetail(400), rChapter(400), rRec(400), rMethod(400)"""

VBS_ADDRESULT_OLD = """Sub AddResult(id, cat, title, status, detail, chapter, rec)
    rID(rCount) = id
    rCat(rCount) = cat
    rTitle(rCount) = title
    rStatus(rCount) = status
    rDetail(rCount) = detail
    rChapter(rCount) = chapter
    rRec(rCount) = rec
    rCount = rCount + 1"""
VBS_ADDRESULT_NEW = """Sub AddResult(id, cat, title, status, detail, chapter, rec)
    ' 评审整改（2026-09-30 第4/8条）：rMethod=验证过程/方法（指导书逐条）；
    ' 核查项名对齐指导书（不一致时以指导书为准，脚本项名以〔脚本项名：…〕并入详情前缀）
    Dim gtitle
    gtitle = GuideTitleOf(id)
    rID(rCount) = id
    rCat(rCount) = cat
    If gtitle <> "" And gtitle <> title Then
        rTitle(rCount) = gtitle
        rDetail(rCount) = "〔脚本项名：" & title & "〕" & detail
    Else
        rTitle(rCount) = title
        rDetail(rCount) = detail
    End If
    rStatus(rCount) = status
    rChapter(rCount) = chapter
    rRec(rCount) = rec
    rMethod(rCount) = MethodOf(id)
    rCount = rCount + 1"""

VBS_JSON_OLD = '''        dataStr = dataStr & "{""ch"":""" & JsonEsc(chPart) & """,""id"":""" & JsonEsc(rID(i)) & """,""cat"":""" & JsonEsc(rCat(i)) & """,""title"":""" & JsonEsc(rTitle(i)) & """,""status"":""" & JsonEsc(rStatus(i)) & """,""detail"":""" & JsonEsc(rDetail(i)) & """,""rec"":""" & JsonEsc(rRec(i)) & """,""guide"":""" & JsonEsc(guidePart) & """}"'''
VBS_JSON_NEW = '''        dataStr = dataStr & "{""ch"":""" & JsonEsc(chPart) & """,""id"":""" & JsonEsc(rID(i)) & """,""cat"":""" & JsonEsc(rCat(i)) & """,""title"":""" & JsonEsc(rTitle(i)) & """,""status"":""" & JsonEsc(rStatus(i)) & """,""method"":""" & JsonEsc(rMethod(i)) & """,""detail"":""" & JsonEsc(rDetail(i)) & """,""rec"":""" & JsonEsc(rRec(i)) & """,""guide"":""" & JsonEsc(guidePart) & """}"'''

VBS_JS_OLD = '''    ts.WriteLine "    tr.innerHTML = ""<td>""+esc(x.ch)+""</td><td class='id'>""+esc(x.id)+""</td><td class='cat'>""+esc(x.cat)+""</td>""+"
    ts.WriteLine "      ""<td class='title'>""+esc(x.title)+""</td>""+"
    ts.WriteLine "      ""<td><span class='badge badge-""+x.status+""'>""+ST[x.status]+""</span></td>""+"
    ts.WriteLine "      ""<td class='detail'>""+esc(x.detail)+""</td>""+"
    ts.WriteLine "      ""<td class='rec'>""+esc(x.rec)+""</td>""+"
    ts.WriteLine "      ""<td class='guide'>""+esc(x.guide)+""</td>"";"'''
VBS_JS_NEW = '''    ts.WriteLine "    tr.innerHTML = ""<td>""+esc(x.ch)+""</td><td class='id'>""+esc(x.id)+""</td><td class='cat'>""+esc(x.cat)+""</td>""+"
    ts.WriteLine "      ""<td class='title'>""+esc(x.title)+""</td>""+"
    ts.WriteLine "      ""<td><span class='badge badge-""+x.status+""'>""+ST[x.status]+""</span></td>""+"
    ts.WriteLine "      ""<td class='method'>""+esc(x.method)+""</td>""+"
    ts.WriteLine "      ""<td class='detail'>""+esc(x.detail)+""</td>""+"
    ts.WriteLine "      ""<td class='rec'>""+(x.status=='fail'?esc(x.rec):"""")+""</td>""+"
    ts.WriteLine "      ""<td class='req'>""+(x.status=='fail'?"""":esc(x.rec))+""</td>""+"
    ts.WriteLine "      ""<td class='guide'>""+esc(x.guide)+""</td>"";"'''

VBS_CSS_OLD = '    ts.WriteLine "td.detail,td.rec{color:#374151;max-width:320px;}"'
VBS_CSS_NEW = '    ts.WriteLine "td.detail,td.rec,td.method,td.req{color:#374151;max-width:320px;}"'

VBS_XLSX_HDR_OLD = '    Dim headers : headers = Array("编号","类别","检查项","状态","详情","修复建议/核查要求","章节")'
VBS_XLSX_HDR_NEW = '    Dim headers : headers = Array("编号","类别","检查项","状态","验证过程/方法","详情","修改建议","安全要求","章节")'
VBS_XLSX_FOR_OLD = """    Dim c
    For c = 1 To 7"""
VBS_XLSX_FOR_NEW = """    Dim c
    For c = 1 To 9"""
VBS_XLSX_CELLS_OLD = """        oWS.Cells(row,5).Value = rDetail(i)
        oWS.Cells(row,6).Value = rRec(i)
        oWS.Cells(row,7).Value = rChapter(i)"""
VBS_XLSX_CELLS_NEW = """        oWS.Cells(row,5).Value = rMethod(i)
        oWS.Cells(row,6).Value = rDetail(i)
        If rStatus(i) = "fail" Then
            oWS.Cells(row,7).Value = rRec(i)
            oWS.Cells(row,8).Value = ""
        Else
            oWS.Cells(row,7).Value = ""
            oWS.Cells(row,8).Value = rRec(i)
        End If
        oWS.Cells(row,9).Value = rChapter(i)"""
VBS_XLSX_WIDTH_OLD = """    oWS.Columns(1).ColumnWidth = 10
    oWS.Columns(2).ColumnWidth = 12
    oWS.Columns(3).ColumnWidth = 40
    oWS.Columns(4).ColumnWidth = 12
    oWS.Columns(5).ColumnWidth = 50
    oWS.Columns(6).ColumnWidth = 50
    oWS.Columns(7).ColumnWidth = 14
    oWS.Columns(5).WrapText = True
    oWS.Columns(6).WrapText = True"""
VBS_XLSX_WIDTH_NEW = """    oWS.Columns(1).ColumnWidth = 10
    oWS.Columns(2).ColumnWidth = 12
    oWS.Columns(3).ColumnWidth = 40
    oWS.Columns(4).ColumnWidth = 12
    oWS.Columns(5).ColumnWidth = 40
    oWS.Columns(6).ColumnWidth = 50
    oWS.Columns(7).ColumnWidth = 50
    oWS.Columns(8).ColumnWidth = 50
    oWS.Columns(9).ColumnWidth = 14
    oWS.Columns(5).WrapText = True
    oWS.Columns(6).WrapText = True
    oWS.Columns(7).WrapText = True
    oWS.Columns(8).WrapText = True"""

VBS_CSV_HDR_OLD = '    ts.WriteLine "编号,类别,检查项,状态,详情,修复建议/核查要求,章节"'
VBS_CSV_HDR_NEW = '    ts.WriteLine "编号,类别,检查项,状态,验证过程/方法,详情,修改建议,安全要求,章节"'
VBS_CSV_ROW_OLD = """        ts.WriteLine CsvEsc(rID(i)) & "," & CsvEsc(rCat(i)) & "," & CsvEsc(rTitle(i)) & "," & _
                    CsvEsc(statusText) & "," & CsvEsc(rDetail(i)) & "," & CsvEsc(rRec(i)) & "," & _
                    CsvEsc(rChapter(i))"""
VBS_CSV_ROW_NEW = """        Dim recFix, reqCol
        If rStatus(i) = "fail" Then
            recFix = rRec(i) : reqCol = ""
        Else
            recFix = "" : reqCol = rRec(i)
        End If
        ts.WriteLine CsvEsc(rID(i)) & "," & CsvEsc(rCat(i)) & "," & CsvEsc(rTitle(i)) & "," & _
                    CsvEsc(statusText) & "," & CsvEsc(rMethod(i)) & "," & CsvEsc(rDetail(i)) & "," & _
                    CsvEsc(recFix) & "," & CsvEsc(reqCol) & "," & CsvEsc(rChapter(i))"""

VBS_XLS_TITLE_OLD = '    ts.WriteLine "<tr><td colspan=""7"" class=""title"">配置核查报告（Windows XP/7版）</td></tr>"'
VBS_XLS_TITLE_NEW = '    ts.WriteLine "<tr><td colspan=""9"" class=""title"">配置核查报告（Windows XP/7版）</td></tr>"'
VBS_XLS_SUB_OLD = '    ts.WriteLine "<tr><td colspan=""7"" class=""sub"">生成时间：" & HtmlEsc(CStr(dtNow)) & "  系统：" & HtmlEsc(osCaption) & "</td></tr>"'
VBS_XLS_SUB_NEW = '    ts.WriteLine "<tr><td colspan=""9"" class=""sub"">生成时间：" & HtmlEsc(CStr(dtNow)) & "  系统：" & HtmlEsc(osCaption) & "</td></tr>"'
VBS_XLS_HDR_OLD = '    ts.WriteLine "<tr><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">详情</th><th class=""hd"">修复建议/核查要求</th><th class=""hd"">章节</th></tr>"'
VBS_XLS_HDR_NEW = '    ts.WriteLine "<tr><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">验证过程/方法</th><th class=""hd"">详情</th><th class=""hd"">修改建议</th><th class=""hd"">安全要求</th><th class=""hd"">章节</th></tr>"'
VBS_XLS_ROW_OLD = '''        ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rTitle(i)) & "</td>" & _
                     "<td class=""" & statusCls & """>" & statusText & "</td>" & _
                     "<td>" & HtmlEsc(rDetail(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rRec(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rChapter(i)) & "</td></tr>"'''
VBS_XLS_ROW_NEW = '''        Dim recFix2, reqCol2
        If rStatus(i) = "fail" Then
            recFix2 = rRec(i) : reqCol2 = ""
        Else
            recFix2 = "" : reqCol2 = rRec(i)
        End If
        ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rTitle(i)) & "</td>" & _
                     "<td class=""" & statusCls & """>" & statusText & "</td>" & _
                     "<td>" & HtmlEsc(rMethod(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rDetail(i)) & "</td>" & _
                     "<td>" & HtmlEsc(recFix2) & "</td>" & _
                     "<td>" & HtmlEsc(reqCol2) & "</td>" & _
                     "<td>" & HtmlEsc(rChapter(i)) & "</td></tr>"'''

ok_vbs = []
if "Function MethodOf(id)" in vbs:
    ok("win/check_xp7.vbs: 已打补丁，跳过")
else:
    for old, new, tag in [(VBS_ARR_OLD, VBS_ARR_NEW, "数组"),
                          (VBS_ADDRESULT_OLD, VBS_ADDRESULT_NEW, "AddResult"),
                          (VBS_JSON_OLD, VBS_JSON_NEW, "JSON字段"),
                          (VBS_JS_OLD, VBS_JS_NEW, "JS表格"),
                          (VBS_CSS_OLD, VBS_CSS_NEW, "CSS"),
                          (VBS_XLSX_HDR_OLD, VBS_XLSX_HDR_NEW, "XLSX表头"),
                          (VBS_XLSX_FOR_OLD, VBS_XLSX_FOR_NEW, "XLSX循环"),
                          (VBS_XLSX_CELLS_OLD, VBS_XLSX_CELLS_NEW, "XLSX单元格"),
                          (VBS_XLSX_WIDTH_OLD, VBS_XLSX_WIDTH_NEW, "XLSX列宽"),
                          (VBS_CSV_HDR_OLD, VBS_CSV_HDR_NEW, "CSV表头"),
                          (VBS_CSV_ROW_OLD, VBS_CSV_ROW_NEW, "CSV行"),
                          (VBS_XLS_TITLE_OLD, VBS_XLS_TITLE_NEW, "XLS标题行"),
                          (VBS_XLS_SUB_OLD, VBS_XLS_SUB_NEW, "XLS子行"),
                          (VBS_XLS_HDR_OLD, VBS_XLS_HDR_NEW, "XLS表头"),
                          (VBS_XLS_ROW_OLD, VBS_XLS_ROW_NEW, "XLS行")]:
        if old in vbs:
            vbs = vbs.replace(old, new, 1)
            ok_vbs.append(tag)
        else:
            ok_vbs.append("!!%s未匹配" % tag)
    # HTML thead（8 列 → 10 列，substring 匹配 ts.WriteLine 行内的表格头）
    n = vbs.count('<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>')
    vbs = vbs.replace('<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>',
                      '<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr></thead>')
    ok_vbs.append("HTML表头×%d" % n)
    # 查询表插到 AddResult 定义之前
    anchor = "Sub AddResult(id, cat, title, status, detail, chapter, rec)"
    if anchor in vbs:
        vbs = vbs.replace(anchor, CASE_BLOCK + "\n" + anchor, 1)
        ok_vbs.append("查询表嵌入")
    else:
        ok_vbs.append("!!查询表锚点未匹配")
    # GBK 可编码性检查（全文），先写临时文件再原子替换，避免编码失败把原文件截断
    try:
        vbs.encode("gbk")
        ok_vbs.append("GBK编码OK")
    except UnicodeEncodeError as e:
        ok("!! check_xp7.vbs 全文含 GBK 外字符：%s" % e)
        raise
    open("win/check_xp7.vbs.tmp", "w", encoding="gbk", newline="").write(vbs)
    os.replace("win/check_xp7.vbs.tmp", "win/check_xp7.vbs")
    ok("win/check_xp7.vbs: %s" % "、".join(ok_vbs))

print("\n".join(report))
