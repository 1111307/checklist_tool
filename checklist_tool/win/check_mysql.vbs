' ============================================================
' 配置核查工具 - MySQL/MariaDB 数据库版（Windows/VBScript）
' 适用系统：Windows XP SP3 / Windows 7（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo check_mysql.vbs（建议管理员权限）
' 连接参数：环境变量 MYSQL_HOST / MYSQL_PORT / MYSQL_USER / MYSQL_PASS
' 输出文件：output\配置核查报告_MySQL_日期时间.html / .xlsx / .xls / .csv
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' --- 结果存储（动态数组）---
Dim rID(60), rCat(60), rTitle(60), rStatus(60)
Dim rDetail(60), rChapter(60), rRec(60)
Dim rCount : rCount = 0

' --- 连接参数 ---
Dim MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASS
MYSQL_HOST = GetSetting("MYSQL_HOST", "127.0.0.1")
MYSQL_PORT = GetSetting("MYSQL_PORT", "3306")
MYSQL_USER = GetSetting("MYSQL_USER", "root")
MYSQL_PASS = GetSetting("MYSQL_PASS", "")

Dim DB_VERSION : DB_VERSION = ""
Dim CONN_OK : CONN_OK = False
Dim MYSQL_PRESENT : MYSQL_PRESENT = False

' ============================================================
' 工具函数
' ============================================================
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

' 多行结果压缩为单行（换行/制表符 -> 空格）
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
' MySQL 查询
' ============================================================
Function MysqlQ(sql)
    Dim cmd
    cmd = "mysql -h" & MYSQL_HOST & " -P" & MYSQL_PORT & " -u" & MYSQL_USER
    If MYSQL_PASS <> "" Then cmd = cmd & " -p" & MYSQL_PASS
    cmd = cmd & " --connect-timeout=5 -N -B -e """ & sql & """ 2>nul"
    MysqlQ = RunCmd(cmd)
End Function

Function GetMysqlVar(key)
    Dim r : r = MysqlQ("SHOW VARIABLES LIKE '" & key & "';")
    Dim lines : lines = Split(r, Chr(10))
    Dim last : last = ""
    Dim i
    For i = 0 To UBound(lines)
        If Trim(lines(i)) <> "" Then last = Trim(lines(i))
    Next
    Dim parts : parts = Split(last, Chr(9))
    If UBound(parts) >= 1 Then
        GetMysqlVar = Trim(parts(1))
    Else
        GetMysqlVar = ""
    End If
End Function

' 版本 EOL 判断：MariaDB <10.7 / MySQL <5.7 视为停止维护
Function MysqlEol(v)
    Dim parts : parts = Split(v, ".")
    Dim maj : maj = 0
    Dim min : min = 0
    If UBound(parts) >= 0 Then
        If IsNumeric(parts(0)) Then maj = CInt(parts(0))
    End If
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(1)) Then min = CInt(parts(1))
    End If
    If InStr(v, "MariaDB") > 0 Then
        If maj = 10 And min <= 6 Then MysqlEol = "eol" Else MysqlEol = "supported"
    Else
        If maj <= 5 Then MysqlEol = "eol" Else MysqlEol = "supported"
    End If
End Function

' ============================================================
' 核查项（共19项，对应配置核查表 MySQL 列）
' ============================================================
Sub Check_1_1()
    If MysqlEol(DB_VERSION) = "eol" Then
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "fail", "当前版本 " & DB_VERSION & " 已停止维护（EOL），无法再获得安全补丁。", "第1章", "升级到仍在维护的版本（MySQL 5.7+/8.0+ / MariaDB 10.11 LTS）。"
    Else
        AddResult "1.1", "系统安全-数据库", "数据库补丁程序", "manual", "当前版本 " & DB_VERSION & "（仍在维护期）。请人工核对是否为最新补丁版本。", "第1章", "及时升级到最新稳定补丁版本。"
    End If
End Sub

Sub Check_1_7()
    Dim anon : anon = Compact(MysqlQ("SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='';"))
    Dim emptyAcc : emptyAcc = Compact(MysqlQ("SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket') AND User NOT IN ('mariadb.sys','mysql.sys','PUBLIC');"))
    If anon <> "" Then
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "fail", "存在匿名账户（无需口令即可登录）：" & anon, "第1章", "删除所有匿名账户，禁止无口令访问。"
    ElseIf emptyAcc <> "" Then
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "fail", "存在空口令账户：" & empty, "第1章", "为所有账户设置强口令。"
    Else
        AddResult "1.7", "系统安全-数据库", "数据库账户管理", "pass", "未发现匿名账户或空口令账户。", "第1章", "持续定期核查账户。"
    End If
End Sub

Sub Check_1_8()
    Dim udf : udf = MysqlQ("SELECT name FROM mysql.func;")
    Dim re : Set re = New RegExp
    re.Pattern = "sys_exec|sys_eval|sys_bineval|sys_get"
    re.IgnoreCase = True
    If re.Test(udf) Then
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "fail", "存在高风险用户自定义函数(UDF)：可调用系统命令（sys_exec/sys_eval 等）。", "第1章", "删除非必要的高危 UDF，禁用 FILE 权限。"
    Else
        AddResult "1.8", "系统安全-数据库", "数据库存储过程管理", "pass", "未发现高风险 UDF。", "第1章", "持续控制存储过程与 UDF 数量。"
    End If
End Sub

Sub Check_1_9()
    Dim danger : danger = Compact(MysqlQ("SELECT CONCAT(u.User,'@',u.Host) FROM mysql.user u WHERE u.User NOT IN ('root','mysql.sys','mysql.session','mysql.infoschema') AND (u.File_priv='Y' OR u.Super_priv='Y' OR u.Process_priv='Y' OR u.Shutdown_priv='Y' OR u.Grant_priv='Y' OR (u.Select_priv='Y' AND u.Insert_priv='Y' AND u.Update_priv='Y' AND u.Delete_priv='Y' AND u.Create_priv='Y' AND u.Drop_priv='Y'));"))
    If danger <> "" Then
        AddResult "1.9", "系统安全-数据库", "数据库权限最小化", "fail", "存在危险全局权限账户：" & danger, "第1章", "收敛账户权限，遵循最小权限原则。"
    Else
        AddResult "1.9", "系统安全-数据库", "数据库权限最小化", "pass", "未发现危险全局权限账户。", "第1章", "持续最小权限。"
    End If
End Sub

Sub Check_1_10()
    Dim bind : bind = GetMysqlVar("bind_address")
    Dim skip : skip = GetMysqlVar("skip_networking")
    If LCase(skip) = "on" Then
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "pass", "skip_networking=ON（仅本地访问）。", "第1章", "保持。"
    ElseIf bind = "127.0.0.1" Or bind = "localhost" Then
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "pass", "bind_address=" & bind & "。", "第1章", "保持。"
    Else
        AddResult "1.10", "系统安全-数据库", "数据库访问控制", "manual", "bind_address=" & bind & "，请人工确认监听范围。", "第1章", "限制监听地址。"
    End If
End Sub

Sub Check_1_11()
    Dim logbin : logbin = GetMysqlVar("log_bin")
    If LCase(logbin) = "on" Then
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "manual", "log_bin=ON，但需人工确认定期备份计划。", "第1章", "建立定期备份并验证可恢复。"
    Else
        AddResult "1.11", "系统安全-数据库", "数据库备份策略", "fail", "log_bin=OFF，无二进制日志/备份机制。", "第1章", "启用 binlog + 定期备份。"
    End If
End Sub

Sub Check_1_12()
    Dim glog : glog = GetMysqlVar("general_log")
    If LCase(glog) = "on" Then
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "pass", "general_log=ON。", "第1章", "升级为专用审计插件。"
    Else
        AddResult "1.12", "系统安全-数据库", "数据库审计插件", "fail", "未启用审计（general_log=OFF）。", "第1章", "启用审计。"
    End If
End Sub

Sub Check_1_13()
    AddResult "1.13", "系统安全-数据库", "数据库与业务隔离", "manual", "数据分类独立存储需人工核查。", "第1章", "分类独立存储。"
End Sub

Sub Check_1_15()
    Dim root_remote : root_remote = Compact(MysqlQ("SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND LOCATE('%', Host) > 0;"))
    Dim bind : bind = GetMysqlVar("bind_address")
    If root_remote <> "" Then
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "fail", "root 允许远程登录：" & root_remote, "第1章", "禁止 root 远程登录。"
    ElseIf bind = "127.0.0.1" Then
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "pass", "已限制（bind=" & bind & "）。", "第1章", "保持。"
    Else
        AddResult "1.15", "系统安全-数据库", "数据库远程访问控制", "manual", "请人工确认超管远程限制。", "第1章", "限制来源。"
    End If
End Sub

Sub Check_1_16()
    AddResult "1.16", "系统安全-数据库", "输入验证/SQL注入防护", "manual", "属应用层能力，需人工核查。", "第1章", "应用层参数化查询。"
End Sub

Sub Check_1_19()
    Dim port : port = GetMysqlVar("port")
    Dim root_empty : root_empty = Compact(MysqlQ("SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket') AND User NOT IN ('mariadb.sys','mysql.sys','PUBLIC');"))
    Dim issues : issues = ""
    If port = "3306" Then issues = "默认端口 3306"
    If root_empty <> "" Then
        If issues <> "" Then issues = issues & "；"
        issues = issues & "root 空口令"
    End If
    If issues <> "" Then
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "fail", issues, "第1章", "改默认端口/口令。"
    Else
        AddResult "1.19", "系统安全-数据库", "数据库默认账户/端口", "pass", "未发现默认端口/空口令。", "第1章", "持续保持。"
    End If
End Sub

Sub Check_1_20()
    Dim vp : vp = GetMysqlVar("validate_password.policy")
    Dim secure : secure = GetMysqlVar("secure_file_priv")
    If vp <> "" Or secure <> "" Then
        AddResult "1.20", "系统安全-数据库", "数据库安全策略", "pass", "检测到安全策略（validate_password=" & vp & "，secure_file_priv=" & secure & "）。", "第1章", "完善策略。"
    Else
        AddResult "1.20", "系统安全-数据库", "数据库安全策略", "fail", "未检测到安全策略。", "第1章", "配置安全策略。"
    End If
End Sub

Sub Check_1_21()
    AddResult "1.21", "系统安全-数据库", "数据库操作审计", "manual", "行/列级审计需人工确认。", "第1章", "配置审计策略。"
End Sub

Sub Check_1_22()
    AddResult "1.22", "系统安全-数据库", "数据库审计日志留存", "manual", "独立监控需人工核查。", "第1章", "部署独立监控。"
End Sub

Sub Check_1_23()
    Dim app_hosts : app_hosts = Compact(MysqlQ("SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE Host='%';"))
    Dim bind : bind = GetMysqlVar("bind_address")
    If app_hosts <> "" Then
        AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "manual", "存在 % 通配账户，请人工确认。", "第1章", "精确限定 Host。"
    ElseIf bind = "127.0.0.1" Then
        AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "pass", "访问受限（bind=" & bind & "）。", "第1章", "保持。"
    Else
        AddResult "1.23", "系统安全-数据库", "应用与数据库账户分离", "manual", "请人工确认。", "第1章", "限制访问。"
    End If
End Sub

Sub Check_1_24()
    Dim expire : expire = GetMysqlVar("expire_logs_days")
    Dim sec : sec = GetMysqlVar("binlog_expire_logs_seconds")
    Dim days : days = 0
    If sec <> "" And IsNumeric(sec) Then
        If CLng(sec) > 0 Then days = CLng(sec) \ 86400
    ElseIf expire <> "" And IsNumeric(expire) Then
        If CLng(expire) > 0 Then days = CLng(expire)
    End If
    If days >= 180 Then
        AddResult "1.24", "系统安全-数据库", "日志审计留存180天", "pass", "binlog 保留约 " & days & " 天。", "第1章", "保持。"
    ElseIf days > 0 Then
        AddResult "1.24", "系统安全-数据库", "日志审计留存180天", "fail", "binlog 保留 " & days & " 天（<180）。", "第1章", "设置 >=180 天。"
    Else
        AddResult "1.24", "系统安全-数据库", "日志审计留存180天", "manual", "请人工核实留存时长。", "第1章", "配置留存。"
    End If
End Sub

Sub Check_1_25()
    Dim fwOut : fwOut = LCase(RunCmd("netsh advfirewall show allprofiles state 2>nul"))
    If InStr(fwOut, "on") = 0 Then fwOut = fwOut & LCase(RunCmd("netsh firewall show opmode 2>nul"))
    If InStr(fwOut, "on") > 0 Or InStr(fwOut, "启用") > 0 Then
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "manual", "检测到防火墙启用，边界防护需人工核查。", "第1章", "部署边界防护。"
    Else
        AddResult "1.25", "系统安全-数据库", "边界保护抗攻击防篡改", "fail", "未检测到防火墙启用。", "第1章", "部署防火墙。"
    End If
End Sub

Sub Check_1_26()
    Dim logerr : logerr = GetMysqlVar("log_error")
    If logerr <> "" And LCase(logerr) <> "stderr" Then
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "pass", "错误日志已配置：" & logerr, "第1章", "保持。"
    Else
        AddResult "1.26", "系统安全-数据库", "防病毒/补丁日志记录", "fail", "未配置错误日志到独立文件。", "第1章", "配置 log_error。"
    End If
End Sub

Sub Check_2_16()
    If MysqlEol(DB_VERSION) = "eol" Then
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "fail", "版本 " & DB_VERSION & " 已 EOL。", "第2章", "升级。"
    Else
        AddResult "2.16", "系统安全-数据库", "补丁修复升级到最新", "manual", "版本 " & DB_VERSION & "，请人工核对最新。", "第2章", "定期升级。"
    End If
End Sub

' ============================================================
' 主流程
' ============================================================
WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - MySQL/MariaDB 版（Windows/VBScript）"
WScript.Echo "  参考标准：配置核查作业指导书v2.2"
WScript.Echo "================================================================"
WScript.Echo "  连接：" & MYSQL_USER & "@" & MYSQL_HOST & ":" & MYSQL_PORT

Dim verOut : verOut = RunCmd("mysql --version 2>nul")
If verOut <> "" Then MYSQL_PRESENT = True

If MYSQL_PRESENT Then
    DB_VERSION = Trim(MysqlQ("SELECT VERSION();"))
    If DB_VERSION <> "" Then
        CONN_OK = True
        WScript.Echo "  数据库版本：" & DB_VERSION
    Else
        WScript.Echo "  [!] 无法连接 MySQL，所有项将标记为需人工核查。"
    End If
Else
    WScript.Echo "  [!] 未检测到 mysql 客户端，所有项将标记为不适用。"
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
    Check_1_20
    Check_1_21
    Check_1_22
    Check_1_23
    Check_1_24
    Check_1_25
    Check_1_26
    Check_2_16
Else
    Dim titles(18, 1)
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
    titles(11,0) = "1.20" : titles(11,1) = "数据库安全策略"
    titles(12,0) = "1.21" : titles(12,1) = "数据库操作审计"
    titles(13,0) = "1.22" : titles(13,1) = "数据库审计日志留存"
    titles(14,0) = "1.23" : titles(14,1) = "应用与数据库账户分离"
    titles(15,0) = "1.24" : titles(15,1) = "日志审计留存180天"
    titles(16,0) = "1.25" : titles(16,1) = "边界保护抗攻击防篡改"
    titles(17,0) = "1.26" : titles(17,1) = "防病毒/补丁日志记录"
    titles(18,0) = "2.16" : titles(18,1) = "补丁修复升级到最新"
    Dim i, st
    For i = 0 To 18
        st = "manual"
        If Not MYSQL_PRESENT Then st = "na"
        AddResult titles(i,0), "系统安全-数据库", titles(i,1), st, "无法连接或未安装 MySQL 客户端。", "第1章", "检查连接参数。"
    Next
End If

WScript.Echo ""
WScript.Echo "================================================================ "
Call PrintSummary()
Call GenerateHTML()
Call GenerateExcel()
WScript.Echo ""
WScript.Echo "核查完成，请到 output\ 目录查看 HTML/XLS 报告。"

' ============================================================
' 汇总
' ============================================================
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

' ============================================================
' 生成 HTML 报告
' ============================================================
Function JsonEsc(s)
    Dim t : t = s
    If IsNull(t) Then t = ""
    t = Replace(t, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    JsonEsc = t
End Function

Function PctOf(a, t)
    If t = 0 Then
        PctOf = "0.0"
    Else
        PctOf = CStr(Round(a * 100 / t, 1))
    End If
End Function

Sub GenerateHTML()
    Dim outDir : outDir = "output"
    If Not oFSO.FolderExists(outDir) Then oFSO.CreateFolder(outDir)
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim ts, reportPath : reportPath = outDir & "\配置核查报告_MySQL_" & stamp & ".html"
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
    Dim p_pass, p_fail, p_manual, p_na
    p_pass = PctOf(nPass, rCount) : p_fail = PctOf(nFail, rCount)
    p_manual = PctOf(nManual, rCount) : p_na = PctOf(nNA, rCount)

    ' DATA 数组（JSON）：rChapter 形如「第2章 第2.3节」，拆为 ch / guide
    Dim dataStr : dataStr = ""
    Dim chPart, guidePart, spPos
    For i = 0 To rCount - 1
        If i > 0 Then dataStr = dataStr & ","
        chPart = CStr(rChapter(i)) : guidePart = ""
        spPos = InStr(chPart, " ")
        If spPos > 0 Then
            guidePart = Trim(Mid(chPart, spPos + 1))
            chPart = Left(chPart, spPos - 1)
        End If
        dataStr = dataStr & "{""ch"":""" & JsonEsc(chPart) & """,""id"":""" & JsonEsc(rID(i)) & """,""cat"":""" & JsonEsc(rCat(i)) & """,""title"":""" & JsonEsc(rTitle(i)) & """,""status"":""" & JsonEsc(rStatus(i)) & """,""detail"":""" & JsonEsc(rDetail(i)) & """,""rec"":""" & JsonEsc(rRec(i)) & """,""guide"":""" & JsonEsc(guidePart) & """}"
    Next

    ts.WriteLine "<style>"
    ts.WriteLine ":root{--bg:#f5f7fa;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#e5e7eb;--brand:#0f3057;--brand2:#1a3c6e;--pass:#15803d;--pass-bg:#ecfdf5;--pass-br:#bbf7d0;--fail:#b91c1c;--fail-bg:#fef2f2;--fail-br:#fecaca;--manual:#b45309;--manual-bg:#fffbeb;--manual-br:#fde68a;--na:#4b5563;--na-bg:#f3f4f6;--na-br:#e5e7eb;}"
    ts.WriteLine "*{box-sizing:border-box;}"
    ts.WriteLine "body{margin:0;font:14px/1.65 ""Segoe UI"",""Microsoft YaHei"",system-ui,sans-serif;color:var(--ink);background:var(--bg);}"
    ts.WriteLine ".wrap{max-width:1240px;margin:0 auto;padding:24px 20px 60px;}"
    ts.WriteLine "header{background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 60%,#2563eb 130%);color:#fff;border-radius:12px;padding:26px 30px;margin-bottom:20px;}"
    ts.WriteLine "header h1{margin:0 0 6px;font-size:22px;letter-spacing:.5px;}"
    ts.WriteLine "header .sub{opacity:.85;font-size:13px;}"
    ts.WriteLine ".meta{display:flex;flex-wrap:wrap;gap:8px 28px;margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.25);font-size:13px;}"
    ts.WriteLine ".meta div{opacity:.95;}"
    ts.WriteLine ".meta b{font-weight:600;opacity:.75;margin-right:6px;}"
    ts.WriteLine ".dash{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}"
    ts.WriteLine ".stat{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 16px 14px;position:relative;overflow:hidden;cursor:pointer;transition:transform .15s, box-shadow .15s;}"
    ts.WriteLine ".stat:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(15,48,87,.10);}"
    ts.WriteLine ".stat .num{font-size:34px;font-weight:700;line-height:1.1;font-variant-numeric:tabular-nums;}"
    ts.WriteLine ".stat .lbl{color:var(--muted);font-size:13px;margin-top:2px;}"
    ts.WriteLine ".stat .bar{height:4px;border-radius:2px;margin-top:12px;background:var(--line);}"
    ts.WriteLine ".stat .bar i{display:block;height:100%;border-radius:2px;}"
    ts.WriteLine ".stat.s-pass .num{color:var(--pass);} .stat.s-pass .bar i{background:var(--pass);}"
    ts.WriteLine ".stat.s-fail .num{color:var(--fail);} .stat.s-fail .bar i{background:var(--fail);}"
    ts.WriteLine ".stat.s-manual .num{color:var(--manual);} .stat.s-manual .bar i{background:var(--manual);}"
    ts.WriteLine ".stat.s-na .num{color:var(--na);} .stat.s-na .bar i{background:var(--na);}"
    ts.WriteLine ".stat .pct{position:absolute;right:14px;top:16px;font-size:12px;color:var(--muted);}"
    ts.WriteLine ".toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:16px;}"
    ts.WriteLine ".toolbar input[type=text]{flex:1;min-width:200px;padding:9px 14px;border:1px solid var(--line);border-radius:8px;font-size:13px;outline:none;background:var(--card);}"
    ts.WriteLine ".toolbar input[type=text]:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.12);}"
    ts.WriteLine ".filters{display:flex;gap:6px;flex-wrap:wrap;}"
    ts.WriteLine ".fbtn{border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:20px;font-size:13px;cursor:pointer;transition:all .15s;}"
    ts.WriteLine ".fbtn:hover{border-color:#2563eb;color:#2563eb;}"
    ts.WriteLine ".fbtn.on{background:var(--brand);border-color:var(--brand);color:#fff;}"
    ts.WriteLine ".count{color:var(--muted);font-size:12px;margin-left:4px;}"
    ts.WriteLine ".panel{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;}"
    ts.WriteLine "table{border-collapse:collapse;width:100%;font-size:13px;}"
    ts.WriteLine "thead th{background:#f8fafc;color:#334155;text-align:left;font-weight:600;padding:10px 12px;border-bottom:2px solid var(--line);white-space:nowrap;position:sticky;top:0;z-index:5;}"
    ts.WriteLine "tbody td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top;}"
    ts.WriteLine "tbody tr:hover{background:#f8fafc;}"
    ts.WriteLine "td.id{font-family:Consolas,monospace;font-weight:600;white-space:nowrap;}"
    ts.WriteLine "td.cat{white-space:nowrap;color:var(--muted);}"
    ts.WriteLine "td.title{min-width:180px;}"
    ts.WriteLine ".badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;white-space:nowrap;border:1px solid;}"
    ts.WriteLine ".badge-pass{color:var(--pass);background:var(--pass-bg);border-color:var(--pass-br);}"
    ts.WriteLine ".badge-fail{color:var(--fail);background:var(--fail-bg);border-color:var(--fail-br);}"
    ts.WriteLine ".badge-manual{color:var(--manual);background:var(--manual-bg);border-color:var(--manual-br);}"
    ts.WriteLine ".badge-na{color:var(--na);background:var(--na-bg);border-color:var(--na-br);}"
    ts.WriteLine "td.detail,td.rec{color:#374151;max-width:320px;}"
    ts.WriteLine ".guide{color:var(--muted);font-size:12px;max-width:260px;}"
    ts.WriteLine ".empty{padding:60px;text-align:center;color:var(--muted);}"
    ts.WriteLine "footer{margin-top:26px;color:var(--muted);font-size:12px;text-align:center;}"
    ts.WriteLine "@media(max-width:900px){.dash{grid-template-columns:repeat(2,1fr);}}"
    ts.WriteLine "@media print{.toolbar{display:none;} .panel{border:none;} body{background:#fff;}}"
    ts.WriteLine "</style>"
    ts.WriteLine "</head>"
    ts.WriteLine "<body>"
    ts.WriteLine "<div class=""wrap"">"
    ts.WriteLine "<header>"
    ts.WriteLine "  <h1>MySQL 数据库配置核查报告</h1>"
    ts.WriteLine "  <div class=""sub"">参考标准：配置核查作业指导书v2.2</div>"
    ts.WriteLine "  <div class=""meta"">"
    ts.WriteLine "    <div>连接：" & HtmlEsc(MYSQL_USER & "@" & MYSQL_HOST & ":" & MYSQL_PORT) & "　版本：" & HtmlEsc(DB_VERSION) & "　核查时间：" & CStr(Now()) & "</div>"
    ts.WriteLine "  </div>"
    ts.WriteLine "</header>"
    ts.WriteLine "<div class=""dash"">"
    ts.WriteLine "  <div class=""stat s-pass"" onclick=""fset('pass')""><div class=""num"">" & nPass & "</div><div class=""lbl"">合规</div><div class=""bar""><i style=""width:" & p_pass & "%""></i></div><div class=""pct"">" & p_pass & "%</div></div>"
    ts.WriteLine "  <div class=""stat s-fail"" onclick=""fset('fail')""><div class=""num"">" & nFail & "</div><div class=""lbl"">不合规</div><div class=""bar""><i style=""width:" & p_fail & "%""></i></div><div class=""pct"">" & p_fail & "%</div></div>"
    ts.WriteLine "  <div class=""stat s-manual"" onclick=""fset('manual')""><div class=""num"">" & nManual & "</div><div class=""lbl"">需人工核查</div><div class=""bar""><i style=""width:" & p_manual & "%""></i></div><div class=""pct"">" & p_manual & "%</div></div>"
    ts.WriteLine "  <div class=""stat s-na"" onclick=""fset('na')""><div class=""num"">" & nNA & "</div><div class=""lbl"">不适用</div><div class=""bar""><i style=""width:" & p_na & "%""></i></div><div class=""pct"">" & p_na & "%</div></div>"
    ts.WriteLine "</div>"
    ts.WriteLine "<div class=""toolbar"">"
    ts.WriteLine "  <input id=""q"" type=""text"" placeholder=""搜索编号 / 核查项 / 详情…"">"
    ts.WriteLine "  <div class=""filters"">"
    ts.WriteLine "    <button class=""fbtn on"" data-f=""all"">全部<span class=""count"">" & rCount & "</span></button>"
    ts.WriteLine "    <button class=""fbtn"" data-f=""fail"">不合规<span class=""count"">" & nFail & "</span></button>"
    ts.WriteLine "    <button class=""fbtn"" data-f=""manual"">需人工<span class=""count"">" & nManual & "</span></button>"
    ts.WriteLine "    <button class=""fbtn"" data-f=""pass"">合规<span class=""count"">" & nPass & "</span></button>"
    ts.WriteLine "    <button class=""fbtn"" data-f=""na"">不适用<span class=""count"">" & nNA & "</span></button>"
    ts.WriteLine "  </div>"
    ts.WriteLine "</div>"
    ts.WriteLine "<div class=""panel"">"
    ts.WriteLine "<table id=""tbl"">"
    ts.WriteLine "<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>"
    ts.WriteLine "<tbody id=""tb""></tbody>"
    ts.WriteLine "</table>"
    ts.WriteLine "<div class=""empty"" id=""empty"" style=""display:none"">没有匹配的核查项</div>"
    ts.WriteLine "</div>"
    ts.WriteLine "<footer>本报告由配置核查工具自动生成 · " & CStr(Now()) & "</footer>"
    ts.WriteLine "</div>"
    ts.WriteLine "<script>"
    ts.WriteLine "var DATA = [" & dataStr & "];"
    ts.WriteLine "var ST = {pass:""合规"", fail:""不合规"", manual:""需人工核查"", na:""不适用""};"
    ts.WriteLine "var curF = ""all"";"
    ts.WriteLine "function esc(s){var d=document.createElement(""div"");d.textContent=s==null?"""":s;return d.innerHTML;}"
    ts.WriteLine "function render(){"
    ts.WriteLine "  var q = document.getElementById(""q"").value.trim().toLowerCase();"
    ts.WriteLine "  var tb = document.getElementById(""tb""); tb.innerHTML = """";"
    ts.WriteLine "  var n = 0;"
    ts.WriteLine "  DATA.forEach(function(x){"
    ts.WriteLine "    if(curF!=""all"" && x.status!=curF) return;"
    ts.WriteLine "    if(q && (x.id+"" ""+x.title+"" ""+x.detail+"" ""+x.cat+"" ""+x.ch).toLowerCase().indexOf(q)<0) return;"
    ts.WriteLine "    n++;"
    ts.WriteLine "    var tr = document.createElement(""tr"");"
    ts.WriteLine "    tr.innerHTML = ""<td>""+esc(x.ch)+""</td><td class='id'>""+esc(x.id)+""</td><td class='cat'>""+esc(x.cat)+""</td>""+"
    ts.WriteLine "      ""<td class='title'>""+esc(x.title)+""</td>""+"
    ts.WriteLine "      ""<td><span class='badge badge-""+x.status+""'>""+ST[x.status]+""</span></td>""+"
    ts.WriteLine "      ""<td class='detail'>""+esc(x.detail)+""</td>""+"
    ts.WriteLine "      ""<td class='rec'>""+esc(x.rec)+""</td>""+"
    ts.WriteLine "      ""<td class='guide'>""+esc(x.guide)+""</td>"";"
    ts.WriteLine "    tb.appendChild(tr);"
    ts.WriteLine "  });"
    ts.WriteLine "  document.getElementById(""empty"").style.display = n? ""none"":""block"";"
    ts.WriteLine "}"
    ts.WriteLine "function fset(f){"
    ts.WriteLine "  curF = f;"
    ts.WriteLine "  var bs = document.querySelectorAll("".fbtn"");"
    ts.WriteLine "  for(var i=0;i<bs.length;i++){ bs[i].className = ""fbtn"" + (bs[i].getAttribute(""data-f"")==f ? "" on"" : """"); }"
    ts.WriteLine "  render();"
    ts.WriteLine "}"
    ts.WriteLine "(function(){"
    ts.WriteLine "  var bs = document.querySelectorAll("".fbtn"");"
    ts.WriteLine "  for(var i=0;i<bs.length;i++){ bs[i].onclick = (function(b){ return function(){ fset(b.getAttribute(""data-f"")); }; })(bs[i]); }"
    ts.WriteLine "})();"
    ts.WriteLine "document.getElementById(""q"").addEventListener(""input"", render);"
    ts.WriteLine "render();"
    ts.WriteLine "</script>"
    ts.WriteLine "</body>"
    ts.WriteLine "</html>"
    ts.Close
    WScript.Echo "[OK] 报告生成完成：" & reportPath
End Sub

' ============================================================
' 生成 Excel 报告（有 Excel 时真 .xlsx；无 Excel 降级 .xls(HTML)）
' ============================================================
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

    oWS.Cells(1,1).Value = "MySQL 数据库配置核查报告"
    oWS.Cells(2,1).Value = "连接：" & MYSQL_USER & "@" & MYSQL_HOST & ":" & MYSQL_PORT & "  版本：" & DB_VERSION
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

    Dim xlsxPath : xlsxPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_MySQL_" & stamp & ".xlsx"
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

' ============================================================
' 生成 .xls 报告（HTML 表格，无 Excel 时兜底）
' ============================================================
Sub GenerateXLS()
    Dim outDir : outDir = "output"
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim ts, xlsPath : xlsPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_MySQL_" & stamp & ".xls"
    Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7""><b>MySQL 数据库配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""7"">连接：" & HtmlEsc(MYSQL_USER & "@" & MYSQL_HOST & ":" & MYSQL_PORT) & "  版本：" & HtmlEsc(DB_VERSION) & "</td></tr>"
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
