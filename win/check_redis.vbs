' ============================================================
' 配置核查工具 - Redis 版（Windows/VBScript）
' 适用系统：Windows XP SP3 / Windows 7（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo check_redis.vbs
' 连接参数：环境变量 REDIS_HOST / REDIS_PORT / REDIS_PASS
' 输出文件：output\配置核查报告_Redis_日期时间.html / .xlsx / .xls
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

Dim rID(60), rCat(60), rTitle(60), rStatus(60)
Dim rDetail(60), rChapter(60), rRec(60)
Dim rCount : rCount = 0

Dim REDIS_HOST, REDIS_PORT, REDIS_PASS
REDIS_HOST = GetSetting("REDIS_HOST", "127.0.0.1")
REDIS_PORT = GetSetting("REDIS_PORT", "6379")
REDIS_PASS = GetSetting("REDIS_PASS", "")

Dim REDIS_VERSION : REDIS_VERSION = ""
Dim CONN_OK : CONN_OK = False
Dim REDIS_PRESENT : REDIS_PRESENT = False

Function ReadConfig(key)
    Dim confFile : confFile = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\db_config.conf"
    If Not oFSO.FileExists(confFile) Then ReadConfig = "" : Exit Function
    Dim f : Set f = oFSO.OpenTextFile(confFile, 1, False, -2)
    Do While Not f.AtEndOfStream
        Dim line : line = Trim(f.ReadLine())
        If line <> "" And Left(line, 1) <> "#" And Left(line, 1) <> ";" Then
            Dim eq : eq = InStr(line, "=")
            If eq > 0 Then
                Dim k : k = Trim(Left(line, eq - 1))
                Dim v : v = Trim(Mid(line, eq + 1))
                If LCase(k) = LCase(key) Then ReadConfig = v : f.Close : Exit Function
            End If
        End If
    Loop
    f.Close
    ReadConfig = ""
End Function

Function GetSetting(key, defaultVal)
    If oShell.Environment("Process")(key) <> "" Then GetSetting = oShell.Environment("Process")(key) : Exit Function
    Dim cfg : cfg = ReadConfig(key)
    If cfg <> "" Then GetSetting = cfg : Exit Function
    GetSetting = defaultVal
End Function

Function RunCmd(cmd)
    On Error Resume Next
    Dim oExec
    Set oExec = oShell.Exec("cmd /c " & cmd)
    Dim output : output = ""
    Do While Not oExec.StdOut.AtEndOfStream
        output = output & oExec.StdOut.ReadLine() & Chr(10)
    Loop
    RunCmd = output
    On Error GoTo 0
End Function

Sub AddResult(id, cat, title, status, detail, chapter, rec)
    rID(rCount) = id
    rCat(rCount) = cat
    rTitle(rCount) = title
    rStatus(rCount) = status
    rDetail(rCount) = detail
    rChapter(rCount) = chapter
    rRec(rCount) = rec
    rCount = rCount + 1
    Dim sym
    Select Case status
        Case "pass":   sym = "[OK] "
        Case "fail":   sym = "[!!] "
        Case "manual": sym = "[??] "
        Case Else:     sym = "[--] "
    End Select
    WScript.Echo "  " & sym & "[" & id & "] " & title
End Sub

Function HtmlEsc(s)
    Dim r : r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, Chr(10), "<br>")
    HtmlEsc = r
End Function

Function Compact(s)
    Dim r : r = s
    r = Replace(r, Chr(13), " ")
    r = Replace(r, Chr(10), " ")
    r = Replace(r, Chr(9), " ")
    Do While InStr(r, "  ") > 0
        r = Replace(r, "  ", " ")
    Loop
    Compact = Trim(r)
End Function

' ============================================================
' Redis 查询
' ============================================================
Function RedisQ(cmd)
    Dim c
    c = "redis-cli -h " & REDIS_HOST & " -p " & REDIS_PORT & " --no-auth-warning"
    If REDIS_PASS <> "" Then c = c & " -a " & REDIS_PASS
    c = c & " " & cmd & " 2>nul"
    RedisQ = RunCmd(c)
End Function

Function GetRedisCfg(key)
    Dim r : r = RedisQ("CONFIG GET " & key)
    Dim lines : lines = Split(r, Chr(10))
    Dim last : last = ""
    Dim i
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then last = Trim(lines(i))
    Next
    GetRedisCfg = last
End Function

Function GetRedisVersion()
    Dim info : info = RedisQ("INFO server")
    Dim re : Set re = New RegExp
    re.Pattern = "redis_version:([0-9.]+)"
    Dim m : Set m = re.Execute(info)
    If m.Count > 0 Then GetRedisVersion = m(0).SubMatches(0)
End Function

' Redis <7 视为 EOL
Function RedisEol(v)
    Dim parts : parts = Split(v, ".")
    Dim maj : maj = 0
    If UBound(parts) >= 0 Then
        If IsNumeric(parts(0)) Then maj = CInt(parts(0))
    End If
    If maj < 7 Then RedisEol = "eol" Else RedisEol = "supported"
End Function

' ============================================================
' 核查项（共16项）
' ============================================================
Sub Check_1_1()
    If RedisEol(REDIS_VERSION) = "eol" Then
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "fail", "Redis 版本 " & REDIS_VERSION & " 已停止维护（<7.0）。", "第1章", "升级到 Redis 7.x/8.x。"
    Else
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "manual", "Redis 版本 " & REDIS_VERSION & "（仍在维护期），请人工核对最新补丁。", "第1章", "及时升级到最新稳定版本。"
    End If
End Sub

Sub Check_1_7()
    Dim pass : pass = GetRedisCfg("requirepass")
    If pass = "" Then
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "fail", "未设置 requirepass，Redis 无需口令即可访问。", "第1章", "设置 requirepass 强口令。"
    Else
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "pass", "已设置 requirepass 认证口令。", "第1章", "持续使用强口令并定期更换。"
    End If
End Sub

Sub Check_1_8()
    Dim risky : risky = ""
    Dim cmds : cmds = Array("FLUSHALL", "FLUSHDB", "EVAL", "SHUTDOWN")
    Dim j
    For j = 0 To UBound(cmds)
        Dim info : info = Replace(Replace(Trim(RedisQ("COMMAND INFO " & cmds(j))), Chr(10), ""), Chr(13), "")
        If info <> "" Then
            If risky <> "" Then risky = risky & " "
            risky = risky & cmds(j)
        End If
    Next
    If risky <> "" Then
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "fail", "以下高危命令未禁用：" & risky, "第1章", "rename-command 禁用高危命令。"
    Else
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "pass", "高危命令已重命名或禁用。", "第1章", "持续禁用不必要的危险命令。"
    End If
End Sub

Sub Check_1_10()
    Dim pm : pm = GetRedisCfg("protected-mode")
    Dim bind : bind = GetRedisCfg("bind")
    If LCase(pm) = "no" And (bind = "" Or InStr(bind, "0.0.0.0") > 0) Then
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "fail", "protected-mode=no 且 bind 为空/0.0.0.0，对全网开放。", "第1章", "开启 protected-mode 并绑定地址。"
    Else
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "pass", "protected-mode=" & pm & "，bind=" & bind & "，访问受限。", "第1章", "持续限制监听地址。"
    End If
End Sub

Sub Check_1_11()
    Dim save : save = GetRedisCfg("save")
    Dim aof : aof = GetRedisCfg("appendonly")
    If save <> "" Then
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "pass", "已配置 RDB 快照（save=" & save & "）。", "第1章", "结合 RDB/AOF 定期备份。"
    ElseIf LCase(aof) = "yes" Then
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "pass", "appendonly=yes。", "第1章", "结合 RDB/AOF 定期备份。"
    Else
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "fail", "未配置持久化（save 空且 appendonly=no）。", "第1章", "开启 RDB/AOF 持久化。"
    End If
End Sub

Sub Check_1_12()
    Dim slow : slow = GetRedisCfg("slowlog-log-slower-than")
    If slow <> "" And slow <> "-1" Then
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "pass", "slowlog 已启用。", "第1章", "结合慢日志与外部审计。"
    Else
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "fail", "slowlog 未启用。", "第1章", "开启 slowlog。"
    End If
End Sub

Sub Check_1_15()
    Dim pass : pass = GetRedisCfg("requirepass")
    Dim bind : bind = GetRedisCfg("bind")
    Dim pm : pm = GetRedisCfg("protected-mode")
    If pass = "" Then
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "fail", "未设置口令，任意来源可远程执行管理命令。", "第1章", "设置口令并限制来源。"
    ElseIf bind = "127.0.0.1" Or LCase(pm) = "yes" Then
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "pass", "已设口令且访问受限。", "第1章", "持续限制远程来源。"
    Else
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "manual", "已设口令但 bind=" & bind & "，请人工确认限制。", "第1章", "通过防火墙/绑定限制来源。"
    End If
End Sub

Sub Check_1_16()
    AddResult "1.16", "系统安全-数据库", "输入验证/SQL注入防护", "manual", "Redis 输入检查属应用层能力，需人工核查。", "第1章", "应用层校验输入。"
End Sub

Sub Check_1_19()
    Dim port : port = GetRedisCfg("port")
    Dim pass : pass = GetRedisCfg("requirepass")
    Dim issues : issues = ""
    If port = "6379" Then issues = "默认端口 6379"
    If pass = "" Then
        If issues <> "" Then issues = issues & "；"
        issues = issues & "未设置口令"
    End If
    If issues <> "" Then
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "fail", issues, "第1章", "改默认端口、设强口令。"
    Else
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "pass", "未发现默认端口/无口令问题。", "第1章", "持续保持。"
    End If
End Sub

Sub Check_1_20()
    Dim maxmem : maxmem = GetRedisCfg("maxmemory")
    Dim pm : pm = GetRedisCfg("protected-mode")
    If (maxmem <> "" And maxmem <> "0") Or LCase(pm) = "yes" Then
        AddResult "1.20", "系统安全-数据库", "数据库安全策略", "pass", "检测到安全策略（maxmemory=" & maxmem & "，protected-mode=" & pm & "）。", "第1章", "完善安全基线。"
    Else
        AddResult "1.20", "系统安全-数据库", "数据库安全策略", "fail", "未检测到明显安全策略。", "第1章", "配置 maxmemory 等策略。"
    End If
End Sub

Sub Check_1_21()
    AddResult "1.21", "系统安全-数据库", "数据库操作审计", "manual", "Redis 无行级/列级审计，需结合外部审计工具。", "第1章", "结合外部审计。"
End Sub

Sub Check_1_22()
    AddResult "1.22", "系统安全-数据库", "数据库审计日志留存", "manual", "独立监控审计需人工核查。", "第1章", "部署独立监控。"
End Sub

Sub Check_1_23()
    Dim bind : bind = GetRedisCfg("bind")
    Dim pm : pm = GetRedisCfg("protected-mode")
    If bind = "127.0.0.1" Or LCase(pm) = "yes" Then
        AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "pass", "访问受限（bind=" & bind & "）。", "第1章", "持续确保仅应用服务器访问。"
    Else
        AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "manual", "请人工确认是否仅应用服务器访问。", "第1章", "绑定内网地址。"
    End If
End Sub

Sub Check_1_25()
    Dim fwOut : fwOut = LCase(RunCmd("netsh advfirewall show allprofiles state 2>nul"))
    If InStr(fwOut, "on") = 0 Then fwOut = fwOut & LCase(RunCmd("netsh firewall show opmode 2>nul"))
    If InStr(fwOut, "on") > 0 Or InStr(fwOut, "启用") > 0 Then
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "manual", "检测到 Windows 防火墙启用，但边界防护需人工核查。", "第1章", "部署边界防护。"
    Else
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "fail", "未检测到防火墙启用。", "第1章", "部署防火墙。"
    End If
End Sub

Sub Check_1_26()
    Dim logfile : logfile = GetRedisCfg("logfile")
    If logfile <> "" And logfile <> """""" Then
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "pass", "已配置日志文件：" & logfile, "第1章", "确保日志完整。"
    Else
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "fail", "logfile 未配置到独立文件。", "第1章", "配置 logfile。"
    End If
End Sub

Sub Check_2_16()
    If RedisEol(REDIS_VERSION) = "eol" Then
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "fail", "Redis 版本 " & REDIS_VERSION & " 已停止维护。", "第2章", "升级到 Redis 7.x/8.x。"
    Else
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "manual", "Redis 版本 " & REDIS_VERSION & "，请人工核对最新。", "第2章", "定期升级。"
    End If
End Sub

' ============================================================
' 主流程
' ============================================================
WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - Redis 版（Windows/VBScript）"
WScript.Echo "  参考标准：配置核查作业指导书正式版2026_4_1"
WScript.Echo "================================================================"
WScript.Echo "  连接：" & REDIS_HOST & ":" & REDIS_PORT

Dim verOut : verOut = RunCmd("redis-cli --version 2>nul")
If verOut <> "" Then REDIS_PRESENT = True

If REDIS_PRESENT Then
    REDIS_VERSION = GetRedisVersion()
    If REDIS_VERSION <> "" Then
        CONN_OK = True
        WScript.Echo "  Redis 版本：" & REDIS_VERSION
    Else
        WScript.Echo "  [!] 无法连接 Redis，所有项将标记为需人工核查。"
    End If
Else
    WScript.Echo "  [!] 未检测到 redis-cli 客户端，所有项将标记为不适用。"
End If
WScript.Echo ""

If CONN_OK Then
    Check_1_1
    Check_1_7
    Check_1_8
    Check_1_10
    Check_1_11
    Check_1_12
    Check_1_15
    Check_1_16
    Check_1_19
    Check_1_20
    Check_1_21
    Check_1_22
    Check_1_23
    Check_1_25
    Check_1_26
    Check_2_16
Else
    Dim titles(15, 1)
    titles(0,0) = "1.1"  : titles(0,1) = "数据库补丁程序"
    titles(1,0) = "1.7"  : titles(1,1) = "数据库账户管理"
    titles(2,0) = "1.8"  : titles(2,1) = "数据库存储过程管理"
    titles(3,0) = "1.10" : titles(3,1) = "数据库访问控制"
    titles(4,0) = "1.11" : titles(4,1) = "数据库备份策略"
    titles(5,0) = "1.12" : titles(5,1) = "数据库审计插件"
    titles(6,0) = "1.15" : titles(6,1) = "数据库远程访问控制"
    titles(7,0) = "1.16" : titles(7,1) = "输入验证/SQL注入防护"
    titles(8,0) = "1.19" : titles(8,1) = "数据库默认账户/端口"
    titles(9,0) = "1.20" : titles(9,1) = "数据库安全策略"
    titles(10,0) = "1.21" : titles(10,1) = "数据库操作审计"
    titles(11,0) = "1.22" : titles(11,1) = "数据库审计日志留存"
    titles(12,0) = "1.23" : titles(12,1) = "应用与数据库账户分离"
    titles(13,0) = "1.25" : titles(13,1) = "边界保护抗攻击防篡改"
    titles(14,0) = "1.26" : titles(14,1) = "防病毒/补丁日志记录"
    titles(15,0) = "2.16" : titles(15,1) = "补丁修复升级到最新"
    Dim i, st
    For i = 0 To 15
        st = "manual"
        If Not REDIS_PRESENT Then st = "na"
        AddResult titles(i,0), "系统安全-数据库", titles(i,1), st, "无法连接或未安装 redis-cli 客户端。", "第1章", "检查连接参数。"
    Next
End If

WScript.Echo ""
WScript.Echo "================================================================ "
Call PrintSummary()
Call GenerateHTML()
Call GenerateExcel()
WScript.Echo ""
WScript.Echo "核查完成，请到 output\ 目录查看 HTML/XLS 报告。"

Sub PrintSummary()
    Dim nPass, nFail, nManual, nNA : nPass = 0 : nFail = 0 : nManual = 0 : nNA = 0
    Dim i
    For i = 0 To rCount - 1
        If rStatus(i) = "pass" Then
            nPass = nPass + 1
        ElseIf rStatus(i) = "fail" Then
            nFail = nFail + 1
        ElseIf rStatus(i) = "manual" Then
            nManual = nManual + 1
        Else
            nNA = nNA + 1
        End If
    Next
    WScript.Echo "  核查完成，共 " & rCount & " 项："
    WScript.Echo "  [OK]  通过       : " & nPass & " 项"
    WScript.Echo "  [!!]  未通过     : " & nFail & " 项"
    WScript.Echo "  [??]  需人工核查 : " & nManual & " 项"
    WScript.Echo "  [--]  不适用     : " & nNA & " 项"
End Sub

Sub GenerateHTML()
    Dim outDir : outDir = "output"
    If Not oFSO.FolderExists(outDir) Then oFSO.CreateFolder(outDir)
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim reportPath : reportPath = outDir & "\配置核查报告_Redis_" & stamp & ".html"
    Dim ts
    Set ts = oFSO.CreateTextFile(reportPath, True, False)

    Dim nPass, nFail, nManual, nNA : nPass = 0 : nFail = 0 : nManual = 0 : nNA = 0
    Dim i
    For i = 0 To rCount - 1
        If rStatus(i) = "pass" Then
            nPass = nPass + 1
        ElseIf rStatus(i) = "fail" Then
            nFail = nFail + 1
        ElseIf rStatus(i) = "manual" Then
            nManual = nManual + 1
        Else
            nNA = nNA + 1
        End If
    Next

    ts.WriteLine "<!DOCTYPE html><html lang=""zh-CN""><head><meta charset=""GBK"">"
    ts.WriteLine "<title>Redis 配置核查报告</title>"
    ts.WriteLine "<style>body{font-family:'Microsoft YaHei',Arial,sans-serif;margin:20px;background:#f5f5f5}"
    ts.WriteLine "h1{color:#333;text-align:center}h2{margin-top:30px}"
    ts.WriteLine ".sum{background:white;padding:20px;margin:20px 0;border-radius:5px}"
    ts.WriteLine ".pass{color:#28a745}.fail{color:#dc3545}.manual{color:#e67e00}.na{color:#6c757d}"
    ts.WriteLine "table{width:100%;border-collapse:collapse;background:white;margin:20px 0}"
    ts.WriteLine "th,td{padding:8px;text-align:left;border-bottom:1px solid #ddd}"
    ts.WriteLine "th{background:#007bff;color:white}"
    ts.WriteLine ".det{white-space:pre-wrap;word-wrap:break-word;max-width:400px}"
    ts.WriteLine ".rec{background:#fff3cd;padding:8px;border-left:3px solid #ffc107}"
    ts.WriteLine ".ch{color:#666;font-size:.9em}</style></head><body>"
    ts.WriteLine "<h1>Redis 配置核查报告</h1>"
    ts.WriteLine "<p style=""text-align:center;color:#666"">连接：" & HtmlEsc(REDIS_HOST & ":" & REDIS_PORT) & "　版本：" & HtmlEsc(REDIS_VERSION) & "　核查时间：" & CStr(Now()) & "</p>"
    ts.WriteLine "<div class=""sum""><h2>核查摘要</h2>"
    ts.WriteLine "<p><strong>总核查项：</strong>" & rCount & " 项</p>"
    ts.WriteLine "<p class=""pass""><strong>通过：</strong>" & nPass & " 项</p>"
    ts.WriteLine "<p class=""fail""><strong>未通过：</strong>" & nFail & " 项</p>"
    ts.WriteLine "<p class=""manual""><strong>需人工核查：</strong>" & nManual & " 项</p>"
    ts.WriteLine "<p class=""na""><strong>不适用：</strong>" & nNA & " 项</p>"
    ts.WriteLine "</div>"

    If nFail > 0 Then
        ts.WriteLine "<h2 class=""fail"">未通过项目（" & nFail & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>详情</th><th>修复建议</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "fail" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td><td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td><td class=""rec"">" & HtmlEsc(rRec(i)) & "</td><td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    If nManual > 0 Then
        ts.WriteLine "<h2 class=""manual"">需人工核查项目（" & nManual & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>核查要求</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "manual" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td><td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td><td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    If nNA > 0 Then
        ts.WriteLine "<h2 class=""na"">不适用项目（" & nNA & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>说明</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "na" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td><td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td><td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    If nPass > 0 Then
        ts.WriteLine "<h2 class=""pass"">通过项目（" & nPass & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>详情</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "pass" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td><td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td><td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    ts.WriteLine "</body></html>"
    ts.Close
    WScript.Echo "[OK] HTML报告已生成：" & reportPath
End Sub

Sub GenerateExcel()
    Dim oExcel
    On Error Resume Next
    Set oExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Call GenerateXLS()
        Exit Sub
    End If
    On Error GoTo 0

    Dim outDir : outDir = "output"
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)

    Dim oWB : Set oWB = oExcel.Workbooks.Add()
    Dim oWS : Set oWS = oWB.Worksheets(1)
    oWS.Name = "核查明细"
    oExcel.Visible = False
    oExcel.DisplayAlerts = False

    oWS.Cells(1,1).Value = "Redis 配置核查报告"
    oWS.Cells(2,1).Value = "连接：" & REDIS_HOST & ":" & REDIS_PORT & "  版本：" & REDIS_VERSION
    oWS.Cells(1,1).Font.Size = 14
    oWS.Cells(1,1).Font.Bold = True

    Dim headers : headers = Array("编号","类别","检查项","状态","详情","修复建议","章节")
    Dim c
    For c = 1 To 7
        oWS.Cells(4,c).Value = headers(c-1)
        oWS.Cells(4,c).Font.Bold = True
        oWS.Cells(4,c).Interior.Color = RGB(0,112,192)
        oWS.Cells(4,c).Font.Color = RGB(255,255,255)
    Next

    Dim i
    For i = 0 To rCount - 1
        Dim row : row = 5 + i
        oWS.Cells(row,1).Value = rID(i)
        oWS.Cells(row,2).Value = rCat(i)
        oWS.Cells(row,3).Value = rTitle(i)
        Dim statusText
        Select Case rStatus(i)
            Case "pass":   statusText = "通过"
            Case "fail":   statusText = "未通过"
            Case "manual": statusText = "需人工核查"
            Case Else:     statusText = "不适用"
        End Select
        oWS.Cells(row,4).Value = statusText
        oWS.Cells(row,5).Value = rDetail(i)
        oWS.Cells(row,6).Value = rRec(i)
        oWS.Cells(row,7).Value = rChapter(i)
        Select Case rStatus(i)
            Case "pass":   oWS.Cells(row,4).Font.Color = RGB(40,167,69)
            Case "fail":   oWS.Cells(row,4).Font.Color = RGB(220,53,69)
            Case "manual": oWS.Cells(row,4).Font.Color = RGB(230,126,0)
        End Select
    Next

    oWS.Columns(1).ColumnWidth = 10
    oWS.Columns(2).ColumnWidth = 14
    oWS.Columns(3).ColumnWidth = 30
    oWS.Columns(4).ColumnWidth = 12
    oWS.Columns(5).ColumnWidth = 50
    oWS.Columns(6).ColumnWidth = 40
    oWS.Columns(7).ColumnWidth = 14
    oWS.Columns(5).WrapText = True
    oWS.Columns(6).WrapText = True

    Dim xlsxPath : xlsxPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Redis_" & stamp & ".xlsx"
    On Error Resume Next
    oWB.SaveAs xlsxPath, 51
    If Err.Number <> 0 Then
        Err.Clear
        oWB.Close False
        oExcel.Quit
        Set oExcel = Nothing
        On Error GoTo 0
        Call GenerateXLS()
        Exit Sub
    End If
    On Error GoTo 0
    oWB.Close False
    oExcel.Quit
    Set oExcel = Nothing
    WScript.Echo "[OK] Excel(.xlsx)报告已生成：" & xlsxPath
End Sub

Sub GenerateXLS()
    Dim outDir : outDir = "output"
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim xlsPath : xlsPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Redis_" & stamp & ".xls"
    Dim ts
    Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7""><b>Redis 配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""7"">连接：" & HtmlEsc(REDIS_HOST & ":" & REDIS_PORT) & "  版本：" & HtmlEsc(REDIS_VERSION) & "</td></tr>"
    ts.WriteLine "<tr><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">详情</th><th class=""hd"">修复建议</th><th class=""hd"">章节</th></tr>"
    Dim i
    For i = 0 To rCount - 1
        Dim statusText
        Select Case rStatus(i)
            Case "pass":   statusText = "通过"
            Case "fail":   statusText = "未通过"
            Case "manual": statusText = "需人工核查"
            Case Else:     statusText = "不适用"
        End Select
        ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td><td>" & HtmlEsc(rTitle(i)) & "</td>"
        ts.WriteLine "<td>" & statusText & "</td><td>" & HtmlEsc(rDetail(i)) & "</td><td>" & HtmlEsc(rRec(i)) & "</td><td>" & HtmlEsc(rChapter(i)) & "</td></tr>"
    Next
    ts.WriteLine "</table></body></html>"
    ts.Close
    WScript.Echo "[OK] Excel(.xls)报告已生成：" & xlsPath
End Sub
