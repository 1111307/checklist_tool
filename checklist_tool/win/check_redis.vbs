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
Dim rDetail(60), rChapter(60), rRec(60), rMethod(60)
Dim rCount : rCount = 0

Dim REDIS_HOST, REDIS_PORT, REDIS_PASS
REDIS_HOST = GetSetting("REDIS_HOST", "127.0.0.1")
REDIS_PORT = GetSetting("REDIS_PORT", "6379")
REDIS_PASS = GetSetting("REDIS_PASS", "")

Dim REDIS_VERSION : REDIS_VERSION = ""
Dim DATA_DIR : DATA_DIR = ""
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

' 验证过程/方法 查询表（评审整改 2026-09-30 第4条；数据源：指导书 v2.0.0 逐条提取）
Function MethodOf(id)
    Select Case id
        Case "1.1": MethodOf = "1.1.1 核查操作系统安装补丁情况：图形界面/文档核查；1.1.2 核查数据库补丁情况：图形界面/文档核查；1.1.3 核查中间件补丁情况：图形界面/文档核查"
        Case "1.7": MethodOf = "1.7.1 Mysql：mysql -uroot -p、SELECT User, Host FROM mysql.user;、INSTALL PLUGIN validate_password SONAME 'validate_password.so';、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc；1.7.2 SQLServer：DROP LOGIN [login_name];、DROP LOGIN [redundant_user];、sqlcmd -S localhost -U SA -P 'SA账户密码'；1.7.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236、ALTER PROFILE DEFAULT LIMIT PASSWORD_LIFE_TIME 90;、ALTER PROFILE DEFAULT LIMIT PASSWORD_GRACE_TIME 7;；1.7.4 Redis：ACL DELUSER redundant_user、redis-cli CONFIG GET requirepass、redis-cli ACL LIST、ps aux | grep redis、systemctl status redis、redis-cli"
        Case "1.8": MethodOf = "1.8.1 Mysql：mysql -uroot -p、DROP PROCEDURE IF EXISTS ‘old_backup_procedure’;、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc、. ~/.bash_profile；1.8.2 SQLServer：USE YourDatabaseName;、sqlcmd -S localhost -U SA -P 'SA账户密码'；1.8.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236、DROP PROCEDURE IF EXISTS 模式名.存储过程名;、DROP PROCEDURE IF EXISTS 存储过程名;；1.8.4 Redis：图形界面/文档核查"
        Case "1.10": MethodOf = "1.10.1 Mysql：mysql -uroot -p、SELECT user, host FROM mysql.user;、SHOW GRANTS FOR 'user'@'host';、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc；1.10.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.10.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.10.4 Redis：CONFIG GET requirepass、CONFIG GET port、CONFIG GET rename-command、redis-cli、CONFIG GET bind、CONFIG GET protected-mode"
        Case "1.11": MethodOf = "1.11.1 Mysql：mysql -uroot -p、mysqladmin -uroot -p create restore_verify、mysql restore_verify < 备份文件路径\dump.sql、mysql < 备份文件路径\dump.sql、USE restore_verify; SHOW TABLES;、vi ~/.bashrc；1.11.2 SQLServer：RESTORE VERIFYONLY FROM DISK = N'备份文件完整路径';、RESTORE FILELISTONLY FROM DISK = N'备份文件完整路径';、sqlcmd -S localhost -U SA -P 'SA账户密码'；1.11.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236、!ls -la /home/dmdba/dmdbms/bin | grep -i backup、./dmrman；1.11.4 Redis：CONFIG GET save、CONFIG GET dir、CONFIG GET *save*、CONFIG GET appendonly、CONFIG GET appendfsync、CONFIG GET *apend*"
        Case "1.12": MethodOf = "1.12.1 Mysql：mysql -uroot -p、SET SESSION sql_safe_updates = 1;、SELECT @@sql_safe_updates;、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc；1.12.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.12.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.12.4 Redis：CONFIG GET rename-command、redis-cli"
        Case "1.15": MethodOf = "1.15.1 MySQL：SELECT user, host FROM mysql.user WHERE user = 'root';、DROP USER 'root'@'%';、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc、. ~/.bash_profile；1.15.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.15.3 达梦：su  -用户名、cd /home/dmdba/dmdbms/data/DAMENG、cat dm.ini | grep -i ""listen""、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.15.4 Redis：redis-cli"
        Case "1.16": MethodOf = "1.16.1 MySQL：mysql -uroot -p、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc、. ~/.bash_profile、env | grep MYSQL；1.16.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.16.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.16.4 Redis：services.msc、ps -ef | grep redis-server | grep -v grep"
        Case "1.19": MethodOf = "1.19.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE 'port';、RENAME USER 'root'@'localhost' TO 'new_admin'@'localhost';FLUSH PRIVILEGES;、ALTER USER '你的管理员用户名'@'主机名' IDENTIFIED BY '你的新复杂密码';FLUSH PRIVILEGES;、vi ~/.bashrc、vi ~/.bash_profile；1.19.2 SQLServer：ALTER LOGIN sa WITH PASSWORD = 'YourNewPassword';、sqlcmd -S localhost -U SA -P 'SA账户密码'；1.19.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.19.4 Redis：图形界面/文档核查"
        Case "1.20": MethodOf = "1.20.1 MySQL：mysql -uroot -p、INSTALL PLUGIN validate_password SONAME 'validate_password.so';、UNINSTALL PLUGIN  validate_password;、GRANT ALL PRIVILEGES ON *.* TO 'km'@'%' IDENTIFIED BY 'km'、vi ~/.bashrc、vi ~/.bash_profile；1.20.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'、sp_configure 'user connections';、sp_configure 'max server memory';、sp_configure 'show advanced options', 1; RECONFIGURE; sp_configure 'xp_cmdshell', 0; RECONFIGURE;、sp_configure 'xp_cmdshell';、ALTER LOGIN 登录名 WITH CHECK_POLICY = ON, CHECK_EXPIRATION = ON;；1.20.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.20.4 Redis：bind 127.0.0.1、protected-mode yes、requirepass ""YourStrongPassword123!""、redis-cli -a ""YourStrongPassword123!"""
        Case "1.21": MethodOf = "1.21.1 MySQL：mysql -uroot -p、SET GLOBAL general_log = 'ON'、SHOW VARIABLES LIKE '%general_log%';、vi ~/.bashrc、vi ~/.bash_profile、source ~/.bashrc；1.21.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.21.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.21.4 Redis：图形界面/文档核查"
        Case "1.22": MethodOf = "1.22.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE '%general_log%';、SET GLOBAL general_log = 'ON';、SET GLOBAL audit_logging = OFF;、SET GLOBAL general_log = OFF;、vi ~/.bashrc；1.22.2 SQLServer：图形界面/文档核查；1.22.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.22.4 Redis：图形界面/文档核查"
        Case "1.23": MethodOf = "1.23.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE 'bind_address';、SHOW GRANTS FOR 'app_user'@'应用服务器IP';、DROP USER '用户名'@'%';、vi ~/.bashrc、vi ~/.bash_profile；1.23.2 SQLServer：图形界面/文档核查；1.23.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录、./disql SYSDBA/SYSDBA@localhost:5236；1.23.4 Redis：图形界面/文档核查"
        Case "1.25": MethodOf = "1.25.1 操作系统抗攻击与防篡改核查：图形界面/文档核查；1.25.2 数据库抗攻击与防篡改核查：图形界面/文档核查"
        Case "1.26": MethodOf = "1.26.1 操作系统安全日志核查：图形界面/文档核查；1.26.2 数据库安全日志核查：图形界面/文档核查"
        Case "2.16": MethodOf = "2.16.1 核查操作系统安装补丁情况：图形界面/文档核查；2.16.2 核查数据库补丁修复和升级到最新版本：图形界面/文档核查；2.16.3 应用软件补丁修复和升级到最新版本：图形界面/文档核查"
        Case Else: MethodOf = ""
    End Select
End Function

' 核查项名 查询表（评审整改 2026-09-30 第8条；指导书 H2 标题）
Function GuideTitleOf(id)
    Select Case id
        Case "1.1": GuideTitleOf = "操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序"
        Case "1.7": GuideTitleOf = "数据库管理系统应删除冗余帐户,应设置不少于8个字符且字符采用字母大小写，数字及特殊字符混合编制的账户口令"
        Case "1.8": GuideTitleOf = "数据库管理系统应删除冗余存储过程"
        Case "1.10": GuideTitleOf = "数据库管理系统应具有自主访问控制功能"
        Case "1.11": GuideTitleOf = "数据库管理系统应具有备份和恢复功能"
        Case "1.12": GuideTitleOf = "数据库管理系统应具有表级审计、告警和阻断功能"
        Case "1.15": GuideTitleOf = "应具备数据库管理系统超级管理员远程登录限制远程登陆限制能力"
        Case "1.16": GuideTitleOf = "应具备数据库管理系统输入（参数）检查能力"
        Case "1.19": GuideTitleOf = "应更换数据库管理系统的默认服务端口、管理员用户名和口令"
        Case "1.20": GuideTitleOf = "数据库管理系统应配置安全策略"
        Case "1.21": GuideTitleOf = "数据库管理系统应具有行级或列级审计功能"
        Case "1.22": GuideTitleOf = "数据库管理系统应采取单独、安全监控、审计措施"
        Case "1.23": GuideTitleOf = "数据库管理系统仅为应用服务器提供访问服务"
        Case "1.25": GuideTitleOf = "检查是否具备边界保护能力,是否可以抗攻击，防纂改"
        Case "1.26": GuideTitleOf = "检查是否有防病毒日志、补丁日志、记录相关信息的完整，有效"
        Case "2.16": GuideTitleOf = "检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本"
        Case Else: GuideTitleOf = ""
    End Select
End Function

Function FixOf(i)
    ' 不合规填修改建议，其余填核查要求（评审 2026-09-30 第5条）
    If rStatus(i) = "fail" Then FixOf = rRec(i) Else FixOf = ""
End Function

Function ReqOf(i)
    If rStatus(i) = "fail" Then ReqOf = "" Else ReqOf = rRec(i)
End Function

Function CleanText(s)
    ' 报告以 GBK 落盘；命令输出里可能混入 UTF-8 字节被逐字节读入的字符（U+0080-U+00FF）
    ' 或控制符/代理对，这些写进 GBK 会产生非法字节并破坏整个 HTML 的编码，统一在此丢弃。
    Dim t : t = s
    If IsNull(t) Then t = ""
    Dim i, ch, code, out
    out = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code = 9 Or code = 10 Or code = 13 Then
            out = out & ch
        ElseIf code < 32 Then
            ' 其他控制符：丢弃
        ElseIf code >= 128 And code <= 255 Then
            ' UTF-8 多字节被当成单字节读入的残留：丢弃
        ElseIf code >= 55296 And code <= 57343 Then
            ' 代理对：丢弃
        Else
            out = out & ch
        End If
    Next
    CleanText = out
End Function

Function Snippet(s, n)
    ' 命令输出只保留可打印 ASCII：中文系统的 netsh/auditpol 输出经管道读入后，
    ' 可能变成 GBK 落盘时不合法的字符，写进报告会破坏整份 HTML 的编码。
    Dim t : t = s
    If IsNull(t) Then t = ""
    Dim i, ch, code, out
    out = ""
    For i = 1 To Len(t)
        ch = Mid(t, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536
        If code = 10 Or code = 13 Or code = 9 Then
            out = out & " "
        ElseIf code >= 32 And code <= 126 Then
            out = out & ch
        End If
    Next
    Do While InStr(out, "  ") > 0
        out = Replace(out, "  ", " ")
    Loop
    out = Trim(out)
    ' 只剩标点/横线这类无信息内容时，也给统一说明，避免详情里出现一串破折号
    Dim hasText : hasText = False
    Dim k, c2
    For k = 1 To Len(out)
        c2 = Mid(out, k, 1)
        If (c2 >= "A" And c2 <= "Z") Or (c2 >= "a" And c2 <= "z") Or (c2 >= "0" And c2 <= "9") Then
            hasText = True
            Exit For
        End If
    Next
    If out = "" Or Not hasText Then
        Snippet = "（原始命令输出已省略：含非 ASCII 字符，可能因编码不一致产生乱码）"
    ElseIf Len(out) > n Then
        Snippet = Left(out, n) & "…"
    Else
        Snippet = out
    End If
End Function

Sub AddResult(id, cat, title, status, detail, chapter, rec)
    title = CleanText(title)
    detail = CleanText(detail)
    rec = CleanText(rec)
    ' 评审整改（2026-09-30 第4/8条）：rMethod=验证过程/方法（指导书逐条）；
    ' 核查项名对齐指导书（不一致时以指导书为准，脚本项名以〔脚本项名：…〕并入详情前缀）
    Dim gtitle : gtitle = GuideTitleOf(id)
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
WScript.Echo "  参考标准：配置核查作业指导书v2.0.0"
WScript.Echo "================================================================"
WScript.Echo "  连接：" & REDIS_HOST & ":" & REDIS_PORT

Dim verOut : verOut = RunCmd("redis-cli --version 2>nul")
If verOut <> "" Then REDIS_PRESENT = True

If REDIS_PRESENT Then
    REDIS_VERSION = GetRedisVersion()
    If REDIS_VERSION <> "" Then
        CONN_OK = True
        DATA_DIR = Trim(GetRedisCfg("dir"))
        If DATA_DIR = "" Then DATA_DIR = "未采集"
        WScript.Echo "  数据目录：" & DATA_DIR
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
    Dim ts, reportPath : reportPath = outDir & "\配置核查报告_Redis_" & stamp & ".html"
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
        dataStr = dataStr & "{""ch"":""" & JsonEsc(chPart) & """,""id"":""" & JsonEsc(rID(i)) & """,""cat"":""" & JsonEsc(rCat(i)) & """,""title"":""" & JsonEsc(rTitle(i)) & """,""status"":""" & JsonEsc(rStatus(i)) & """,""detail"":""" & JsonEsc(rDetail(i)) & """,""method"":""" & JsonEsc(rMethod(i)) & """,""detail"":""" & JsonEsc(rDetail(i)) & """,""fix"":""" & JsonEsc(FixOf(i)) & """,""req"":""" & JsonEsc(ReqOf(i)) & """,""guide"":""" & JsonEsc(guidePart) & """}"
    Next

    ts.WriteLine "<!DOCTYPE html>"
    ts.WriteLine "<html lang=""zh-CN"">"
    ts.WriteLine "<head>"
    ts.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">"
    ts.WriteLine "<title>Redis 配置核查报告</title>"
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
    ts.WriteLine "td.method,td.detail,td.fix,td.req{color:#374151;max-width:320px;}"
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
    ts.WriteLine "  <h1>Redis 配置核查报告</h1>"
    ts.WriteLine "  <div class=""sub"">参考标准：配置核查作业指导书v2.0.0</div>"
    ts.WriteLine "  <div class=""meta"">"
    ts.WriteLine "    <div>连接：" & HtmlEsc(REDIS_HOST & ":" & REDIS_PORT) & "　版本：" & HtmlEsc(REDIS_VERSION) & "　数据目录：" & HtmlEsc(DATA_DIR) & "　核查时间：" & CStr(Now()) & "</div>"
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
    ts.WriteLine "<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>核查要求</th><th>参考指导书</th></tr></thead>"
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
    ts.WriteLine "      ""<td class='method'>""+esc(x.method)+""</td>""+"
    ts.WriteLine "      ""<td class='detail'>""+esc(x.detail)+""</td>""+"
    ts.WriteLine "      ""<td class='fix'>""+esc(x.fix)+""</td>""+"
    ts.WriteLine "      ""<td class='req'>""+esc(x.req)+""</td>""+"
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
    oWS.Cells(2,1).Value = "连接：" & REDIS_HOST & ":" & REDIS_PORT & "  版本：" & REDIS_VERSION & "  数据目录：" & DATA_DIR
    oWS.Cells(1,1).Font.Size = 14
    oWS.Cells(1,1).Font.Bold = True

    Dim headers : headers = Array("编号","类别","检查项","状态","验证过程/方法","详情","修改建议","核查要求","章节")
    Dim c
    For c = 1 To 9
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
        oWS.Cells(row,5).Value = rMethod(i)
        oWS.Cells(row,6).Value = rDetail(i)
        oWS.Cells(row,7).Value = FixOf(i)
        oWS.Cells(row,8).Value = ReqOf(i)
        oWS.Cells(row,9).Value = rChapter(i)
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
    oWS.Columns(6).ColumnWidth = 50
    oWS.Columns(7).ColumnWidth = 30
    oWS.Columns(8).ColumnWidth = 30
    oWS.Columns(9).ColumnWidth = 14
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
    ' 汇总合并（merge_report.vbs）只认 .xls，装了 Excel 时也要落一份
    Call GenerateXLS()
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
    ts.WriteLine "<tr><td colspan=""9""><b>Redis 配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""9"">连接：" & HtmlEsc(REDIS_HOST & ":" & REDIS_PORT) & "  版本：" & HtmlEsc(REDIS_VERSION) & "  数据目录：" & HtmlEsc(DATA_DIR) & "</td></tr>"
    ts.WriteLine "<tr><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">验证过程/方法</th><th class=""hd"">详情</th><th class=""hd"">修改建议</th><th class=""hd"">核查要求</th><th class=""hd"">章节</th></tr>"
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
        ts.WriteLine "<td>" & statusText & "</td><td>" & HtmlEsc(rMethod(i)) & "</td><td>" & HtmlEsc(rDetail(i)) & "</td><td>" & HtmlEsc(FixOf(i)) & "</td><td>" & HtmlEsc(ReqOf(i)) & "</td><td>" & HtmlEsc(rChapter(i)) & "</td></tr>"
    Next
    ts.WriteLine "</table></body></html>"
    ts.Close
    WScript.Echo "[OK] Excel(.xls)报告已生成：" & xlsPath
End Sub
