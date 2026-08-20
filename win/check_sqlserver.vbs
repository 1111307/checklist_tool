' ============================================================
' 配置核查工具 - SQL Server 版（Windows/VBScript）
' 适用系统：Windows XP SP3 / Windows 7（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo check_sqlserver.vbs
' 连接参数：环境变量 MSSQL_HOST / MSSQL_PORT / MSSQL_USER / MSSQL_PASS
' 输出文件：output\配置核查报告_SQLServer_日期时间.html / .xlsx / .xls
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

Dim rID(60), rCat(60), rTitle(60), rStatus(60)
Dim rDetail(60), rChapter(60), rRec(60)
Dim rCount : rCount = 0

Dim MSSQL_HOST, MSSQL_PORT, MSSQL_USER, MSSQL_PASS
MSSQL_HOST = GetSetting("MSSQL_HOST", "localhost")
MSSQL_PORT = GetSetting("MSSQL_PORT", "1433")
MSSQL_USER = GetSetting("MSSQL_USER", "sa")
MSSQL_PASS = GetSetting("MSSQL_PASS", "")

Dim DB_VERSION : DB_VERSION = ""
Dim CONN_OK : CONN_OK = False
Dim SQLCMD_PRESENT : SQLCMD_PRESENT = False

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

' ============================================================
' SQL Server 查询
' ============================================================
Function InvokeSql(sql)
    Dim conn : conn = MSSQL_HOST & "," & MSSQL_PORT
    Dim cmd : cmd = "sqlcmd -S " & conn & " -U " & MSSQL_USER & " -P " & MSSQL_PASS & " -C -h -1 -W -Q ""SET NOCOUNT ON; " & sql & """ 2>nul"
    InvokeSql = Trim(RunCmd(cmd))
End Function

Function SqlMaj(v)
    Dim parts : parts = Split(v, ".")
    If UBound(parts) >= 0 Then
        If IsNumeric(parts(0)) Then SqlMaj = CInt(parts(0))
    End If
End Function

Function SqlBuild(v)
    Dim parts : parts = Split(v, ".")
    If UBound(parts) >= 2 Then
        If IsNumeric(parts(2)) Then SqlBuild = CInt(parts(2))
    End If
End Function

' SQL Server <2017（maj<14）视为 EOL
Function SqlEol(v)
    If SqlMaj(v) < 14 Then SqlEol = "eol" Else SqlEol = "supported"
End Function

' RTM 无累积更新（build<4000）
Function SqlCuStatus(v)
    If SqlBuild(v) < 4000 Then SqlCuStatus = "rtm" Else SqlCuStatus = "cu"
End Function

' ============================================================
' 核查项（共18项）
' ============================================================
Sub Check_1_1()
    If SqlEol(DB_VERSION) = "eol" Then
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "fail", "版本 " & DB_VERSION & " 已 EOL（<2017）。", "第1章", "升级。"
    ElseIf SqlCuStatus(DB_VERSION) = "rtm" Then
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "fail", "版本 " & DB_VERSION & " 为 RTM（未安装累积更新 CU）。", "第1章", "安装最新累积更新。"
    Else
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "manual", "版本 " & DB_VERSION & "，请人工核对是否最新 CU。", "第1章", "安装最新累积更新。"
    End If
End Sub

Sub Check_1_7()
    Dim sa_policy : sa_policy = InvokeSql("SELECT is_policy_checked FROM sys.sql_logins WHERE name='sa';")
    If sa_policy = "0" Then
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "fail", "sa 账户未启用密码策略。", "第1章", "启用密码策略。"
    Else
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "pass", "sa 已启用密码策略。", "第1章", "保持。"
    End If
End Sub

Sub Check_1_8()
    Dim xp : xp = InvokeSql("SELECT value_in_use FROM sys.configurations WHERE name='xp_cmdshell';")
    If xp = "1" Then
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "fail", "xp_cmdshell 已启用。", "第1章", "禁用 xp_cmdshell。"
    Else
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "pass", "xp_cmdshell 已禁用。", "第1章", "保持。"
    End If
End Sub

Sub Check_1_9()
    Dim obj : obj = InvokeSql("SELECT COUNT(*) FROM sys.database_permissions WHERE class_desc IN ('OBJECT_OR_COLUMN','SCHEMA');")
    If obj <> "" And CInt(obj) > 0 Then
        AddResult "1.9", "系统安全-数据库", "数据库权限最小化", "pass", "存在对象级细粒度授权 " & obj & " 条。", "第1章", "持续最小权限。"
    Else
        AddResult "1.9", "系统安全-数据库", "数据库权限最小化", "manual", "请人工确认权限最小化。", "第1章", "最小权限授权。"
    End If
End Sub

Sub Check_1_10()
    Dim cnt : cnt = InvokeSql("SELECT COUNT(*) FROM sys.server_role_members rm JOIN sys.server_principals p ON rm.member_principal_id=p.principal_id JOIN sys.server_principals r ON rm.role_principal_id=r.principal_id WHERE r.name='sysadmin' AND p.name NOT IN ('sa');")
    If cnt <> "" And CInt(cnt) > 0 Then
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "manual", "存在 " & cnt & " 个非 sa 的 sysadmin，请人工确认。", "第1章", "收敛 sysadmin。"
    Else
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "pass", "sysadmin 仅 sa。", "第1章", "保持。"
    End If
End Sub

Sub Check_1_11()
    Dim days : days = InvokeSql("SELECT ISNULL(DATEDIFF(day, MAX(backup_finish_date), GETDATE()), 9999) FROM msdb.dbo.backupset WHERE type='D';")
    If days = "" Or days = "9999" Then
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "fail", "未发现全量备份记录。", "第1章", "建立定期备份。"
    ElseIf CInt(days) <= 7 Then
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "pass", "最近备份距今 " & days & " 天。", "第1章", "保持。"
    Else
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "fail", "最近备份距今 " & days & " 天（>7）。", "第1章", "建立定期备份。"
    End If
End Sub

Sub Check_1_12()
    Dim audit : audit = InvokeSql("SELECT COUNT(*) FROM sys.server_audits WHERE is_state_enabled=1;")
    If audit <> "" And CInt(audit) > 0 Then
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "pass", "已启用服务器审计 " & audit & " 项。", "第1章", "保持。"
    Else
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "fail", "未启用服务器审计。", "第1章", "启用审计。"
    End If
End Sub

Sub Check_1_13()
    AddResult "1.13", "系统安全-数据库", "数据库与业务隔离", "manual", "数据分类独立存储需人工核查。", "第1章", "分类独立存储。"
End Sub

Sub Check_1_15()
    Dim sa_dis : sa_dis = InvokeSql("SELECT is_disabled FROM sys.sql_logins WHERE name='sa';")
    If sa_dis = "1" Then
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "pass", "sa 已禁用。", "第1章", "保持。"
    Else
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "manual", "sa 启用中，请人工确认远程限制。", "第1章", "禁用 sa 或限制远程。"
    End If
End Sub

Sub Check_1_16()
    AddResult "1.16", "系统安全-数据库", "输入验证/SQL注入防护", "manual", "属应用层能力，需人工核查。", "第1章", "应用层参数化查询。"
End Sub

Sub Check_1_19()
    Dim sa_en : sa_en = InvokeSql("SELECT CASE WHEN is_disabled=0 THEN 1 ELSE 0 END FROM sys.sql_logins WHERE name='sa';")
    Dim issues : issues = ""
    If MSSQL_PORT = "1433" Then issues = "默认端口 1433"
    If sa_en = "1" Then
        If issues <> "" Then issues = issues & "；"
        issues = issues & "sa 账户仍启用"
    End If
    If issues <> "" Then
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "fail", issues, "第1章", "改默认端口、加固 sa。"
    Else
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "pass", "未发现默认问题。", "第1章", "保持。"
    End If
End Sub

Sub Check_1_21()
    Dim spec : spec = InvokeSql("SELECT COUNT(*) FROM sys.server_audit_specifications;")
    If spec <> "" And CInt(spec) > 0 Then
        AddResult "1.21", "系统安全-数据库", "数据库操作审计", "manual", "检测到审计规范 " & spec & " 项，行/列级审计粒度需人工确认。", "第1章", "配置审计规范覆盖敏感操作。"
    Else
        AddResult "1.21", "系统安全-数据库", "数据库操作审计", "fail", "未配置审计规范，不具备行/列级审计。", "第1章", "创建审计规范。"
    End If
End Sub

Sub Check_1_22()
    AddResult "1.22", "系统安全-数据库", "数据库审计日志留存", "manual", "独立监控需人工核查。", "第1章", "部署独立监控。"
End Sub

Sub Check_1_23()
    AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "manual", "请人工确认是否仅应用服务器访问。", "第1章", "限制访问。"
End Sub

Sub Check_1_24()
    AddResult "1.24", "系统安全-数据库", "日志审计留存180天", "manual", "审计日志留存时长需人工核查。", "第1章", "配置留存180天。"
End Sub

Sub Check_1_25()
    Dim fwOut : fwOut = LCase(RunCmd("netsh advfirewall show allprofiles state 2>nul"))
    If InStr(fwOut, "on") = 0 Then fwOut = fwOut & LCase(RunCmd("netsh firewall show opmode 2>nul"))
    If InStr(fwOut, "on") > 0 Or InStr(fwOut, "启用") > 0 Then
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "manual", "检测到 Windows 防火墙已启用，边界防护细节需人工核查。", "第1章", "部署完整边界防护。"
    Else
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "fail", "Windows 防火墙未启用，缺乏边界保护。", "第1章", "启用防火墙。"
    End If
End Sub

Sub Check_1_26()
    Dim log : log = InvokeSql("SELECT COUNT(*) FROM sys.dm_os_server_diagnostics_log_configurations;")
    If log <> "" And CInt(log) > 0 Then
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "pass", "检测到诊断日志配置。", "第1章", "保持。"
    Else
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "manual", "请人工核查日志完整性。", "第1章", "配置日志。"
    End If
End Sub

Sub Check_2_16()
    If SqlEol(DB_VERSION) = "eol" Then
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "fail", "版本 " & DB_VERSION & " 已 EOL。", "第2章", "升级。"
    ElseIf SqlCuStatus(DB_VERSION) = "rtm" Then
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "fail", "版本 " & DB_VERSION & " 为 RTM（未安装累积更新 CU）。", "第2章", "安装最新累积更新。"
    Else
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "manual", "版本 " & DB_VERSION & "，请人工核对最新。", "第2章", "定期升级。"
    End If
End Sub

' ============================================================
' 主流程
' ============================================================
WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - SQL Server 版（Windows/VBScript）"
WScript.Echo "  参考标准：配置核查作业指导书正式版2026_4_1"
WScript.Echo "================================================================"
WScript.Echo "  连接：" & MSSQL_USER & "@" & MSSQL_HOST & ":" & MSSQL_PORT

Dim sqlcmdCheck : sqlcmdCheck = RunCmd("sqlcmd -? 2>&1")
If Trim(sqlcmdCheck) <> "" Then SQLCMD_PRESENT = True

If Not SQLCMD_PRESENT Then
    DB_VERSION = "未安装sqlcmd"
    WScript.Echo "  [!] 未检测到 sqlcmd 客户端，所有项将标记为不适用。"
Else
    DB_VERSION = InvokeSql("SELECT SERVERPROPERTY('ProductVersion');")
    If DB_VERSION <> "" Then
        CONN_OK = True
        WScript.Echo "  SQL Server 版本：" & DB_VERSION
    Else
        WScript.Echo "  [!] 无法连接 SQL Server，所有项将标记为需人工核查。"
    End If
End If
WScript.Echo ""

If CONN_OK Then
    Check_1_1
    Check_1_7
    Check_1_8
    Check_1_9
    Check_1_10
    Check_1_11
    Check_1_12
    Check_1_13
    Check_1_15
    Check_1_16
    Check_1_19
    Check_1_21
    Check_1_22
    Check_1_23
    Check_1_24
    Check_1_25
    Check_1_26
    Check_2_16
Else
    Dim titles(17, 1)
    titles(0,0) = "1.1"  : titles(0,1) = "数据库补丁程序"
    titles(1,0) = "1.7"  : titles(1,1) = "数据库账户管理"
    titles(2,0) = "1.8"  : titles(2,1) = "数据库存储过程管理"
    titles(3,0) = "1.9"  : titles(3,1) = "数据库权限最小化"
    titles(4,0) = "1.10" : titles(4,1) = "数据库访问控制"
    titles(5,0) = "1.11" : titles(5,1) = "数据库备份策略"
    titles(6,0) = "1.12" : titles(6,1) = "数据库审计插件"
    titles(7,0) = "1.13" : titles(7,1) = "数据库与业务隔离"
    titles(8,0) = "1.15" : titles(8,1) = "数据库远程访问控制"
    titles(9,0) = "1.16" : titles(9,1) = "输入验证/SQL注入防护"
    titles(10,0) = "1.19" : titles(10,1) = "数据库默认账户/端口"
    titles(11,0) = "1.21" : titles(11,1) = "数据库操作审计"
    titles(12,0) = "1.22" : titles(12,1) = "数据库审计日志留存"
    titles(13,0) = "1.23" : titles(13,1) = "应用与数据库账户分离"
    titles(14,0) = "1.24" : titles(14,1) = "日志审计留存180天"
    titles(15,0) = "1.25" : titles(15,1) = "边界保护抗攻击防篡改"
    titles(16,0) = "1.26" : titles(16,1) = "防病毒/补丁日志记录"
    titles(17,0) = "2.16" : titles(17,1) = "补丁修复升级到最新"
    Dim i, st
    For i = 0 To 17
        st = "manual"
        If Not SQLCMD_PRESENT Then st = "na"
        AddResult titles(i,0), "系统安全-数据库", titles(i,1), st, "无法连接或未安装 sqlcmd。", "第1章", "检查连接参数。"
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
    Dim reportPath : reportPath = outDir & "\配置核查报告_SQLServer_" & stamp & ".html"
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
    ts.WriteLine "<title>SQL Server 配置核查报告</title>"
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
    ts.WriteLine "<h1>SQL Server 数据库配置核查报告</h1>"
    ts.WriteLine "<p style=""text-align:center;color:#666"">连接：" & HtmlEsc(MSSQL_USER & "@" & MSSQL_HOST & ":" & MSSQL_PORT) & "　版本：" & HtmlEsc(DB_VERSION) & "　核查时间：" & CStr(Now()) & "</p>"
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

    oWS.Cells(1,1).Value = "SQL Server 数据库配置核查报告"
    oWS.Cells(2,1).Value = "连接：" & MSSQL_USER & "@" & MSSQL_HOST & ":" & MSSQL_PORT & "  版本：" & DB_VERSION
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

    Dim xlsxPath : xlsxPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_SQLServer_" & stamp & ".xlsx"
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
    Dim xlsPath : xlsPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_SQLServer_" & stamp & ".xls"
    Dim ts
    Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7""><b>SQL Server 数据库配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""7"">连接：" & HtmlEsc(MSSQL_USER & "@" & MSSQL_HOST & ":" & MSSQL_PORT) & "  版本：" & HtmlEsc(DB_VERSION) & "</td></tr>"
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
