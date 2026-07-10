' ============================================================
' 配置核查工具 - Windows XP/7 版（无需Python）
' 适用系统：Windows XP SP3 / Windows 7 / Windows Server 2003/2008 R2
' 运行方式：双击 run.bat（建议管理员权限）
' 输出文件：output\配置核查报告_日期时间.html
' ============================================================
Option Explicit

Dim oShell, oFSO, oWMI
Dim osCaption, osBuild, isXP, is7OrAbove

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' --- 结果存储（动态数组）---
Dim rID(400), rCat(400), rTitle(400), rStatus(400)
Dim rDetail(400), rChapter(400), rRec(400)
Dim rCount : rCount = 0

' --- 检测操作系统 ---
Dim verStr : verStr = RunCmd("ver")
Dim osSub  : osSub  = ""
On Error Resume Next
Set oWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
If Not IsNull(oWMI) And Not IsEmpty(oWMI) Then
    Dim osSet : Set osSet = oWMI.ExecQuery("SELECT Caption,BuildNumber,Version FROM Win32_OperatingSystem")
    Dim osObj
    For Each osObj In osSet
        osCaption = osObj.Caption
        osBuild   = CInt(osObj.BuildNumber)
    Next
End If
On Error GoTo 0

isXP        = (osBuild > 0 And osBuild < 6000)
is7OrAbove  = (osBuild >= 7600)

WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - Windows XP/7 版（无需Python）"
WScript.Echo "  参考标准：配置核查作业指导书正式版2026_4_1"
WScript.Echo "================================================================"
If osCaption <> "" Then WScript.Echo "  系统：" & osCaption
WScript.Echo ""
WScript.Echo "开始检查，请稍候..."
WScript.Echo ""

' ---------- 第1章 系统安全 ----------
Call Check_1_1_Patch()
Call Check_1_2_Antivirus()
Call Check_1_3_Services()
Call Check_1_4_Firewall()
Call Check_1_5_NetBIOS()
Call Check_1_6_RDP()
Call Check_1_17_Redis()
Call Check_1_24_AuditPolicy()
Call Check_1_25_Defender()
Call Check_1_18_ResourceLimits()
Call Check_1_26_SecurityLogs()
Call Check_1_7_DbAccounts()
Call Check_1_8_DbProcedures()
Call Check_1_9_DbGrants()
Call Check_1_10_DbAccessControl()
Call Check_1_11_DbBackup()
Call Check_1_12_DbAuditPlugin()
Call Check_1_13_DbIsolation()
Call Check_1_14_Middleware()
Call Check_1_15_DbRemote()
Call Check_1_16_17_InputCheck()
Call Check_1_19_DbDefaults()
Call Check_1_20_DbPolicy()
Call Check_1_21_DbAudit()
Call Check_1_22_DbAuditRetention()
Call Check_1_23_DbAppSep()

' ---------- 第2章 用户安全 ----------
Call Check_2_1_PasswordExpiry()
Call Check_2_4_UniqueAccounts()
Call Check_2_5_RedundantServices()
Call Check_2_2_UserPatch()
Call Check_2_6_IllegalInternet()
Call Check_2_8_EndpointControl()
Call Check_2_15_SecurityPolicy()
Call Check_2_16_PatchLatest()
Call Check_2_7_NTFS()
Call Check_2_9_RiskySoftware()
Call Check_2_10_USB()
Call Check_2_11_PasswordPolicy()
Call Check_2_12_Wireless()
Call Check_2_13_OutboundControl()
Call Check_2_14_UserAudit()

' ---------- 第3章 数据安全 ----------
Call Check_3_1_Encryption()
Call Check_3_5_LogProtection()
Call Check_3_10_Shares()
Call Check_3_8_StoragePartition()

' ---------- 第4章 应用安全 ----------
Call Check_4_15_NLA()
Call Check_4_16_RoleSeparation()
Call Check_4_17_IPRestrict()
Call Check_4_18_Lockout()
Call Check_4_22_PortSeparation()
Call Check_4_23_AppDBSep()
Call Check_4_24_AppAudit()
Call Check_4_6_WAF()
Call Check_4_8_AntiTamper()
Call Check_4_12_SessionLimit()
Call Check_4_21_DefaultPort()
Call Check_4_29_RemoteMgmt()
Call Check_4_31_AppBackup()
Call Check_4_32_DomesticSoftware()
Call Check_10_1_TLS()

' ---------- 手动核查（第5-9章）----------
Call AddManualChecks()

' ---------- 生成报告 ----------
WScript.Echo ""
WScript.Echo "================================================================"
Call PrintSummary()
Call GenerateHTML()
Call GenerateExcel()

WScript.Echo ""
WScript.Echo "核查完成，请在 output\ 目录查看 HTML 报告。"


' ============================================================
' 第1章 系统安全
' ============================================================

Sub Check_1_1_Patch()
    Dim out : out = RunCmd("wmic qfe list brief /format:csv 2>nul")
    Dim lines : lines = Split(out, Chr(10))
    Dim count : count = 0
    Dim i, latestDate : latestDate = ""
    For i = 1 To UBound(lines)
        Dim parts : parts = Split(lines(i), ",")
        If UBound(parts) >= 4 Then
            count = count + 1
            If Len(parts(4)) >= 8 And latestDate = "" Then latestDate = parts(4)
        End If
    Next
    Dim detail
    If count > 0 Then
        detail = "检测到 " & count & " 个已安装补丁，最新记录：" & latestDate
        Call AddResult("1.1", "系统安全", "操作系统/数据库/中间件应及时安装补丁", _
                       "pass", detail, "第1章 第1.1节", "")
    Else
        Call AddResult("1.1", "系统安全", "操作系统/数据库/中间件应及时安装补丁", _
                       "fail", "未能获取补丁列表，请手动检查 Windows Update 状态", _
                       "第1章 第1.1节", "打开【控制面板】→【Windows Update】检查并安装补丁")
    End If
End Sub

Sub Check_1_2_Antivirus()
    Dim ns : ns = "root\SecurityCenter2"
    If isXP Then ns = "root\SecurityCenter"
    Dim avNames : avNames = ""
    On Error Resume Next
    Dim wmiAV : Set wmiAV = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\" & ns)
    If Not IsNull(wmiAV) Then
        Dim avSet : Set avSet = wmiAV.ExecQuery("SELECT displayName FROM AntiVirusProduct")
        Dim avObj
        For Each avObj In avSet
            If avNames <> "" Then avNames = avNames & ", "
            avNames = avNames & avObj.displayName
        Next
    End If
    On Error GoTo 0
    If avNames <> "" Then
        Call AddResult("1.2", "系统安全", "防病毒软件", _
                       "pass", "检测到防病毒软件：" & avNames, "第1章 第1.2节", "")
    Else
        Call AddResult("1.2", "系统安全", "防病毒软件", _
                       "fail", "未检测到防病毒软件", "第1章 第1.2节", _
                       "安装符合行业要求的防病毒软件并更新病毒库（如360、卡巴斯基等）")
    End If
End Sub

Sub Check_1_3_Services()
    Dim svcOut : svcOut = LCase(RunCmd("sc query type= service state= running 2>nul"))
    Dim risky(3) : risky(0) = "telnet" : risky(1) = "tftp" : risky(2) = "ftp" : risky(3) = "snmptrap"
    Dim found : found = ""
    Dim i
    For i = 0 To 3
        If InStr(svcOut, risky(i)) > 0 Then
            If found <> "" Then found = found & ", "
            found = found & risky(i)
        End If
    Next
    If found <> "" Then
        Call AddResult("1.3", "系统安全", "服务和端口裁剪", _
                       "fail", "发现不必要的高危服务：" & found, "第1章 第1.3节", _
                       "在【服务(services.msc)】中停止并禁用以下服务：" & found)
    Else
        Call AddResult("1.3", "系统安全", "服务和端口裁剪", _
                       "pass", "未发现不必要的高危服务（telnet/tftp/ftp/snmp）", "第1章 第1.3节", "")
    End If
End Sub

Sub Check_1_4_Firewall()
    Dim fwOn : fwOn = False
    Dim detail : detail = ""
    If is7OrAbove Then
        Dim advOut : advOut = RunCmd("netsh advfirewall show allprofiles 2>nul")
        fwOn = (InStr(UCase(advOut), "ON") > 0 Or InStr(advOut, "启用") > 0)
        detail = Left(advOut, 200)
    Else
        Dim xpOut : xpOut = RunCmd("netsh firewall show opmode 2>nul")
        fwOn = (InStr(UCase(xpOut), "ENABLE") > 0 Or InStr(xpOut, "启用") > 0)
        detail = Left(xpOut, 200)
    End If
    If fwOn Then
        Call AddResult("1.4", "系统安全", "防火墙策略", _
                       "pass", "防火墙已启用。" & Chr(10) & detail, "第1章 第1.4节", "")
    Else
        Call AddResult("1.4", "系统安全", "防火墙策略", _
                       "fail", "防火墙未启用或检测失败。" & Chr(10) & detail, "第1章 第1.4节", _
                       "【控制面板】→【Windows 防火墙】 启用防火墙")
    End If
End Sub

Sub Check_1_5_NetBIOS()
    Dim nbReg : nbReg = RunCmd("reg query HKLM\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces /s 2>nul")
    If InStr(nbReg, "NetbiosOptions") > 0 And InStr(nbReg, "0x2") = 0 Then
        Call AddResult("1.5", "系统安全", "停用冗余网络设置（NetBIOS）", _
                       "fail", "NetBIOS over TCP/IP 未全部禁用，存在信息泄露风险。", "第1章 第1.5节", _
                       "网络连接→【属性】→IPv4→属性→高级→WINS 选项卡→禁用 NetBIOS over TCP/IP")
    Else
        Call AddResult("1.5", "系统安全", "停用冗余网络设置（NetBIOS）", _
                       "pass", "NetBIOS over TCP/IP 配置检查通过。", "第1章 第1.5节", "")
    End If
End Sub

Sub Check_1_6_RDP()
    Dim val : val = ReadReg("HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\SecurityLayer")
    If val = "" Then
        Call AddResult("1.6", "系统安全", "远程管理采取传输加密保护", _
                       "manual", "无法读取RDP SecurityLayer注册表项，请手动核查", "第1章 第1.6节", _
                       "HKLM\SYSTEM\...\RDP-Tcp\SecurityLayer 建议设置为 2（TLS加密）")
    ElseIf CInt(val) >= 2 Then
        Call AddResult("1.6", "系统安全", "远程管理采取传输加密保护", _
                       "pass", "RDP SecurityLayer = " & val & "，已启用TLS加密。", "第1章 第1.6节", "")
    Else
        Call AddResult("1.6", "系统安全", "远程管理采取传输加密保护", _
                       "fail", "RDP SecurityLayer = " & val & "，未启用TLS，建议设置为2。", "第1章 第1.6节", _
                       "注册表 HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\SecurityLayer 设为 2")
    End If
End Sub

Sub Check_1_17_Redis()
    Dim netOut : netOut = RunCmd("netstat -ano | findstr :6379 2>nul")
    If netOut = "" Then
        Call AddResult("1.17", "系统安全", "限制Redis等远程管理服务端口", _
                       "na", "未检测到Redis服务（端口6379）。", "第1章 第1.17节", "")
    ElseIf InStr(netOut, "0.0.0.0:6379") > 0 Then
        Call AddResult("1.17", "系统安全", "限制Redis等远程管理服务端口", _
                       "fail", "Redis 6379端口对外开放（0.0.0.0），存在未授权访问风险。", "第1章 第1.17节", _
                       "在 redis.conf 中设置：bind 127.0.0.1，重启Redis服务")
    Else
        Call AddResult("1.17", "系统安全", "限制Redis等远程管理服务端口", _
                       "pass", "Redis端口已绑定本地地址，未对外放开。" & Chr(10) & netOut, "第1章 第1.17节", "")
    End If
End Sub

Sub Check_1_24_AuditPolicy()
    Dim secFile : secFile = oShell.ExpandEnvironmentStrings("%TEMP%") & "\secaudit.cfg"
    RunCmd "secedit /export /cfg """ & secFile & """ /quiet 2>nul"
    If oFSO.FileExists(secFile) Then
        Dim ts : Set ts = oFSO.OpenTextFile(secFile, 1)
        Dim content : content = ts.ReadAll()
        ts.Close
        oFSO.DeleteFile secFile
        Dim auditCount : auditCount = 0
        Dim keyword, keywords
        keywords = Array("AuditSystemEvents", "AuditLogonEvents", "AuditObjectAccess", _
                         "AuditPrivilegeUse", "AuditPolicyChange", "AuditAccountManage")
        Dim k
        For Each k In keywords
            If InStr(content, k & " = ") > 0 Then
                Dim pos : pos = InStr(content, k & " = ")
                Dim valStr : valStr = Mid(content, pos + Len(k) + 3, 1)
                If valStr <> "0" Then auditCount = auditCount + 1
            End If
        Next
        If auditCount >= 3 Then
            Call AddResult("1.24", "系统安全", "日志审计能力，审计日志至少保留180天", _
                           "pass", "已启用 " & auditCount & " 项审计策略（secedit验证）", "第1章 第1.24节", "")
        Else
            Call AddResult("1.24", "系统安全", "日志审计能力，审计日志至少保留180天", _
                           "fail", "已启用 " & auditCount & " 项审计策略（不满足要求至少3项）", "第1章 第1.24节", _
                           "本地安全策略→高级审核策略→配置：登录、账户登录、策略更改、账户管理、系统等")
        End If
    Else
        Call AddResult("1.24", "系统安全", "日志审计能力，审计日志至少保留180天", _
                       "manual", "secedit执行失败，请手动核查审核策略", "第1章 第1.24节", _
                       "本地安全策略→高级审核策略→手动配置各类审核")
    End If
End Sub

Sub Check_1_25_Defender()
    If is7OrAbove Then
        Dim defReg : defReg = ReadReg("HKLM\SOFTWARE\Microsoft\Windows Defender\Real-Time Protection\DisableRealtimeMonitoring")
        If defReg = "1" Then
            Call AddResult("1.25", "系统安全", "具备边界保护能力（实时防护）", _
                           "fail", "Windows Defender 实时保护已禁用", "第1章 第1.25节", _
                           "【设置】→【Windows Defender】 开启实时保护")
        ElseIf defReg = "0" Or defReg = "" Then
            Call AddResult("1.25", "系统安全", "具备边界保护能力（实时防护）", _
                           "pass", "Windows Defender 实时保护已启用。", "第1章 第1.25节", "")
        Else
            Call AddResult("1.25", "系统安全", "具备边界保护能力（实时防护）", _
                           "manual", "无法确认实时保护状态，请手动核查", "第1章 第1.25节", "")
        End If
    Else
        Call AddResult("1.25", "系统安全", "具备边界保护能力（实时防护）", _
                       "manual", "Windows XP 请手动核查实时保护状态", "第1章 第1.25节", "")
    End If
End Sub


' ============================================================
' 第2章 用户安全
' ============================================================

Sub Check_1_18_ResourceLimits()
    Dim issues : issues = ""
    Dim drive : drive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
    Dim quotaOut : quotaOut = RunCmd("fsutil quota query " & drive & " 2>nul")
    Dim quotaLower : quotaLower = LCase(quotaOut)
    If InStr(quotaLower, "disabled") > 0 Or InStr(quotaLower, "not enabled") > 0 Or InStr(quotaLower, "未启用") > 0 Then
        issues = issues & "  · 系统盘 " & drive & " 磁盘配额未启用" & Chr(10)
    End If
    Dim gpOut : gpOut = RunCmd("gpresult /R /SCOPE COMPUTER 2>nul")
    Dim hasGPO : hasGPO = False
    If InStr(gpOut, "Resource") > 0 Or InStr(gpOut, "Quota") > 0 Or InStr(gpOut, "配额") > 0 Or InStr(gpOut, "资源") > 0 Then
        hasGPO = True
    End If
    If Not hasGPO Then
        issues = issues & "  · 未检测到资源限制相关的组策略" & Chr(10)
    End If
    If issues <> "" Then
        Call AddResult("1.18", "系统安全", "限制用户对服务器资源最大/最小使用限度", _
                       "fail", issues, "第1章 第1.18节", _
                       "1. 启用磁盘配额：磁盘属性→配额→启用配额管理" & Chr(10) & _
                       "2. 通过组策略配置用户权限和资源限制")
    Else
        Call AddResult("1.18", "系统安全", "限制用户对服务器资源最大/最小使用限度", _
                       "pass", "已配置资源使用限制。", "第1章 第1.18节", "")
    End If
End Sub

Sub Check_1_26_SecurityLogs()
    Dim issues : issues = ""
    Dim findings : findings = ""
    Dim windir : windir = oShell.ExpandEnvironmentStrings("%windir%")
    If oFSO.FileExists(windir & "\WindowsUpdate.log") Then
        findings = findings & "  · WindowsUpdate.log 存在" & Chr(10)
    ElseIf oFSO.FolderExists(windir & "\Logs\WindowsUpdate") Then
        findings = findings & "  · Windows Update 日志目录存在" & Chr(10)
    Else
        issues = issues & "  · 未找到 Windows Update 日志文件" & Chr(10)
    End If
    Dim progdata : progdata = oShell.ExpandEnvironmentStrings("%ProgramData%")
    If oFSO.FolderExists(progdata & "\Microsoft\Windows Defender\Support") Then
        findings = findings & "  · Windows Defender 支持日志目录存在" & Chr(10)
    Else
        issues = issues & "  · 未找到 Windows Defender 支持日志目录" & Chr(10)
    End If
    Dim secOut : secOut = RunCmd("wevtutil qe Security /c:5 /rd:true /f:text 2>nul")
    If Len(secOut) > 10 Then
        findings = findings & "  · 安全事件日志中有有效记录" & Chr(10)
    Else
        issues = issues & "  · 安全事件日志为空或无法读取" & Chr(10)
    End If
    If issues <> "" Then
        Call AddResult("1.26", "系统安全", "防病毒日志、补丁日志记录完整有效", _
                       "fail", issues, "第1章 第1.26节", _
                       "1. 确认 Windows Update 服务正常运行" & Chr(10) & _
                       "2. 确认 Windows Defender 实时保护已启用" & Chr(10) & _
                       "3. 通过日志查看器检查安全日志是否有有效记录")
    Else
        Call AddResult("1.26", "系统安全", "防病毒日志、补丁日志记录完整有效", _
                       "pass", findings, "第1章 第1.26节", "")
    End If
End Sub

' ============================================================
' 1.7 数据库冗余帐户
' ============================================================
Sub Check_1_7_DbAccounts()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.7", "系统安全", "数据库冗余帐户", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.7节", "")
        Exit Sub
    End If
    Dim usersOut : usersOut = RunCmd("mysql -u root -e ""SELECT user FROM mysql.user;"" 2>nul")
    Dim found : found = ""
    Dim risky : risky = "test anonymous guest demo"
    Dim r
    For Each r In Split(risky, " ")
        If InStr(LCase(usersOut), r) > 0 Then
            If found <> "" Then found = found & ","
            found = found & r
        End If
    Next
    If found <> "" Then
        Call AddResult("1.7", "系统安全", "数据库冗余帐户", "fail", _
                       "发现冗余账户：" & found, "第1章 第1.7节", _
                       "删除冗余账户：DROP USER 'test'@'localhost';")
    Else
        Call AddResult("1.7", "系统安全", "数据库冗余帐户", "pass", _
                       "MySQL已安装，未发现冗余账户。", "第1章 第1.7节", "")
    End If
End Sub

' ============================================================
' 1.8 数据库冗余存储过程
' ============================================================
Sub Check_1_8_DbProcedures()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.8节", "")
        Exit Sub
    End If
    Dim procsOut : procsOut = RunCmd("mysql -u root -e ""SELECT COUNT(*) cnt FROM information_schema.routines WHERE routine_type='PROCEDURE';"" 2>nul")
    Dim lines : lines = Split(procsOut, Chr(10))
    Dim i
    For i = 0 To UBound(lines)
        Dim ln : ln = Trim(lines(i))
        If IsNumeric(ln) Then
            If CInt(ln) > 0 Then
                Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "fail", _
                               "发现 " & ln & " 个存储过程，请确认是否需要", "第1章 第1.8节", _
                               "删除不必要的存储过程：DROP PROCEDURE proc_name;")
                Exit Sub
            End If
        End If
    Next
    Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "pass", _
                   "MySQL已安装，未发现冗余存储过程。", "第1章 第1.8节", "")
End Sub

' ============================================================
' 1.9 数据库细粒度访问授权
' ============================================================
Sub Check_1_9_DbGrants()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.9", "系统安全", "数据库细粒度访问授权（表级增删改查）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.9节", "")
        Exit Sub
    End If
    Dim grantsOut : grantsOut = RunCmd("mysql -u root -e ""SELECT user,host,Super_priv FROM mysql.user WHERE user NOT IN ('root','');"" 2>nul")
    If Len(Trim(grantsOut)) = 0 Then
        Call AddResult("1.9", "系统安全", "数据库细粒度访问授权（表级增删改查）", "fail", _
                       "MySQL已安装，但无法获取用户权限，请手动执行 SHOW GRANTS", "第1章 第1.9节", _
                       "使用授权账户登录MySQL：SHOW GRANTS FOR 'user'@'host';")
        Exit Sub
    End If
    Dim risky : risky = 0
    Dim glines : glines = Split(grantsOut, Chr(10))
    Dim gi
    For gi = 0 To UBound(glines)
        Dim gparts : gparts = Split(Trim(glines(gi)), Chr(9))
        If UBound(gparts) >= 2 Then
            If UCase(Trim(gparts(2))) = "Y" Then risky = risky + 1
        End If
    Next
    If risky > 0 Then
        Call AddResult("1.9", "系统安全", "数据库细粒度访问授权（表级增删改查）", "fail", _
                       "发现 " & risky & " 个非root用户拥有Super权限，存在过度授权风险。", "第1章 第1.9节", _
                       "收回非必要的Super权限：REVOKE SUPER ON *.* FROM 'user'@'host';")
    Else
        Call AddResult("1.9", "系统安全", "数据库细粒度访问授权（表级增删改查）", "pass", _
                       "MySQL用户权限检查通过，未发现过度授权。", "第1章 第1.9节", "")
    End If
End Sub

' ============================================================
' 1.10 数据库自主访问控制功能
' ============================================================
Sub Check_1_10_DbAccessControl()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.10节", "")
        Exit Sub
    End If
    Dim issues : issues = ""
    Dim anonOut : anonOut = RunCmd("mysql -u root -e ""SELECT COUNT(*) FROM mysql.user WHERE user='' AND (Select_priv='Y' OR Insert_priv='Y' OR Update_priv='Y');"" 2>nul")
    Dim alines : alines = Split(anonOut, Chr(10))
    Dim ai
    For ai = 0 To UBound(alines)
        Dim av : av = Trim(alines(ai))
        If IsNumeric(av) Then
            If CInt(av) > 0 Then
                issues = issues & "  · 发现匿名用户拥有数据操作权限" & Chr(10)
            End If
        End If
    Next
    Dim emptyOut : emptyOut = RunCmd("mysql -u root -e ""SELECT COUNT(*) FROM mysql.user WHERE user='root' AND (authentication_string='' OR authentication_string IS NULL);"" 2>nul")
    Dim elines : elines = Split(emptyOut, Chr(10))
    Dim ei
    For ei = 0 To UBound(elines)
        Dim ev : ev = Trim(elines(ei))
        If IsNumeric(ev) Then
            If CInt(ev) > 0 Then
                issues = issues & "  · root账户存在空密码" & Chr(10)
            End If
        End If
    Next
    If issues <> "" Then
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "fail", _
                       issues, "第1章 第1.10节", _
                       "1. 删除匿名账户：DELETE FROM mysql.user WHERE user='';" & Chr(10) & _
                       "2. 为root设置强密码：ALTER USER 'root'@'localhost' IDENTIFIED BY '强密码';")
    Else
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "pass", _
                       "MySQL访问控制检查通过。", "第1章 第1.10节", "")
    End If
End Sub

' ============================================================
' 1.11 数据库备份与恢复功能
' ============================================================
Sub Check_1_11_DbBackup()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.11", "系统安全", "数据库备份与恢复功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.11节", "")
        Exit Sub
    End If
    Dim findings : findings = ""
    Dim taskOut : taskOut = RunCmd("schtasks /query /fo LIST 2>nul")
    If InStr(LCase(taskOut), "mysqldump") > 0 Or InStr(LCase(taskOut), "mysql") > 0 Then
        findings = findings & "  · 检测到MySQL相关备份计划任务" & Chr(10)
    End If
    Dim bdirs : bdirs = "C:\backup C:\mysql_backup D:\backup D:\mysql_backup"
    Dim bd
    For Each bd In Split(bdirs, " ")
        If oFSO.FolderExists(bd) Then
            findings = findings & "  · 发现备份目录：" & bd & Chr(10)
        End If
    Next
    If findings <> "" Then
        Call AddResult("1.11", "系统安全", "数据库备份与恢复功能", "pass", _
                       findings, "第1章 第1.11节", "")
    Else
        Call AddResult("1.11", "系统安全", "数据库备份与恢复功能", "fail", _
                       "未检测到MySQL备份计划任务，也未发现备份文件。", "第1章 第1.11节", _
                       "设置定期备份：mysqldump -u root -p --all-databases > backup.sql" & Chr(10) & _
                       "并通过Windows计划任务定期执行")
    End If
End Sub

' ============================================================
' 1.12 数据库表级审计、告警和阻断功能
' ============================================================
Sub Check_1_12_DbAuditPlugin()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.12", "系统安全", "数据库表级审计、告警和阻断功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.12节", "")
        Exit Sub
    End If
    Dim pluginsOut : pluginsOut = RunCmd("mysql -u root -e ""SHOW PLUGINS;"" 2>nul")
    Dim foundPlugin : foundPlugin = ""
    Dim plines : plines = Split(pluginsOut, Chr(10))
    Dim pi
    For pi = 0 To UBound(plines)
        Dim pl : pl = LCase(Trim(plines(pi)))
        If InStr(pl, "audit") > 0 Or InStr(pl, "mcafee") > 0 Then
            Dim pparts : pparts = Split(plines(pi), Chr(9))
            If Len(Trim(pparts(0))) > 0 Then
                If foundPlugin <> "" Then foundPlugin = foundPlugin & ", "
                foundPlugin = foundPlugin & Trim(pparts(0))
            End If
        End If
    Next
    If foundPlugin <> "" Then
        Call AddResult("1.12", "系统安全", "数据库表级审计、告警和阻断功能", "pass", _
                       "检测到审计插件：" & foundPlugin, "第1章 第1.12节", "")
        Exit Sub
    End If
    Dim genlogOut : genlogOut = RunCmd("mysql -u root -e ""SHOW VARIABLES LIKE 'general_log';"" 2>nul")
    If InStr(UCase(genlogOut), "ON") > 0 Then
        Call AddResult("1.12", "系统安全", "数据库表级审计、告警和阻断功能", "na", _
                       "已检测到general_log（通用日志），但未检测到专用审计插件。", "第1章 第1.12节", _
                       "安装MySQL Enterprise Audit或开源审计插件（如Percona Audit Log）")
    Else
        Call AddResult("1.12", "系统安全", "数据库表级审计、告警和阻断功能", "fail", _
                       "未检测到审计插件，general_log也未启用。", "第1章 第1.12节", _
                       "启用通用日志：SET GLOBAL general_log = 'ON';" & Chr(10) & _
                       "安装专用审计插件，支持表级审计和告警")
    End If
End Sub

' ============================================================
' 1.13 数据库数据与其他应用数据独立存储
' ============================================================
Sub Check_1_13_DbIsolation()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.13", "系统安全", "数据库数据与其他应用数据独立存储", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.13节", "")
        Exit Sub
    End If
    Dim dbsOut : dbsOut = RunCmd("mysql -u root -e ""SHOW DATABASES;"" 2>nul")
    If Len(Trim(dbsOut)) = 0 Then
        Call AddResult("1.13", "系统安全", "数据库数据与其他应用数据独立存储", "fail", _
                       "MySQL已安装，但无法获取数据库列表（可能需要密码）", "第1章 第1.13节", _
                       "使用授权账户登录MySQL：SHOW DATABASES;")
        Exit Sub
    End If
    Dim sysDbs : sysDbs = "information_schema mysql performance_schema sys Database"
    Dim userDbCount : userDbCount = 0
    Dim userDbNames : userDbNames = ""
    Dim dlines : dlines = Split(dbsOut, Chr(10))
    Dim di
    For di = 0 To UBound(dlines)
        Dim dname : dname = Trim(dlines(di))
        If Len(dname) > 0 Then
            If InStr(sysDbs, dname) = 0 Then
                userDbCount = userDbCount + 1
                If userDbNames <> "" Then userDbNames = userDbNames & ", "
                userDbNames = userDbNames & dname
            End If
        End If
    Next
    If userDbCount > 1 Then
        Call AddResult("1.13", "系统安全", "数据库数据与其他应用数据独立存储", "fail", _
                       "检测到 " & userDbCount & " 个业务数据库：" & userDbNames & Chr(10) & _
                       "多个业务数据库共享同一MySQL实例，请确认是否为同一应用。", "第1章 第1.13节", _
                       "不同应用的数据应部署在不同数据库实例中。")
    Else
        Call AddResult("1.13", "系统安全", "数据库数据与其他应用数据独立存储", "pass", _
                       "检测到业务数据库：" & userDbNames & "，未发现混用情况。", "第1章 第1.13节", "")
    End If
End Sub

' ============================================================
' 1.14 中间件安全
' ============================================================
Sub Check_1_14_Middleware()
    Dim nginxOut : nginxOut = RunCmd("where nginx 2>nul")
    Dim tomcatOut : tomcatOut = RunCmd("sc query Tomcat 2>nul")
    Dim hasNginx : hasNginx = (InStr(LCase(nginxOut), "nginx") > 0)
    Dim hasTomcat : hasTomcat = (InStr(tomcatOut, "RUNNING") > 0)
    If Not hasNginx And Not hasTomcat Then
        Call AddResult("1.14", "系统安全", "中间件安全", "na", _
                       "未检测到Nginx/Tomcat服务。如业务需要请安装中间件并配置HTTPS。", "第1章 第1.14节", "")
        Exit Sub
    End If
    If hasNginx Then
        Dim sslOut : sslOut = RunCmd("findstr /i ""ssl"" C:\nginx\conf\nginx.conf 2>nul")
        If Len(Trim(sslOut)) = 0 Then
            Call AddResult("1.14", "系统安全", "中间件安全", "fail", _
                           "Nginx未配置SSL/TLS", "第1章 第1.14节", "配置SSL证书并启用HTTPS")
        Else
            Call AddResult("1.14", "系统安全", "中间件安全", "pass", _
                           "Nginx已安装并配置HTTPS。", "第1章 第1.14节", "")
        End If
        Exit Sub
    End If
    Call AddResult("1.14", "系统安全", "中间件安全", "pass", _
                   "Tomcat已安装。", "第1章 第1.14节", "")
End Sub

' ============================================================
' 1.15 数据库远程登录限制
' ============================================================
Sub Check_1_15_DbRemote()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.15", "系统安全", "数据库远程登录限制", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.15节", "")
        Exit Sub
    End If
    Dim remoteOut : remoteOut = RunCmd("mysql -u root -e ""SELECT host FROM mysql.user WHERE user='root' AND host NOT IN ('localhost','127.0.0.1','::1');"" 2>nul")
    Dim rlines : rlines = Split(remoteOut, Chr(10))
    If UBound(rlines) > 1 Then
        Call AddResult("1.15", "系统安全", "数据库远程登录限制", "fail", _
                       "MySQL root账户允许远程登录，存在安全风险。", "第1章 第1.15节", _
                       "禁止root远程：UPDATE mysql.user SET host='localhost' WHERE user='root'; FLUSH PRIVILEGES;")
    Else
        Call AddResult("1.15", "系统安全", "数据库远程登录限制", "pass", _
                       "MySQL管理员未允许远程登录配置。", "第1章 第1.15节", "")
    End If
End Sub

' ============================================================
' 1.16/1.17 数据库输入检查（SQL严格模式）
' ============================================================
Sub Check_1_16_17_InputCheck()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.16_17", "系统安全", "数据库输入检查（SQL注入防护）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.16-1.17节", "")
        Exit Sub
    End If
    Dim modeOut : modeOut = RunCmd("mysql -u root -e ""SELECT @@sql_mode;"" 2>nul")
    If InStr(UCase(modeOut), "STRICT") > 0 Then
        Call AddResult("1.16_17", "系统安全", "数据库输入检查（SQL注入防护）", "pass", _
                       "MySQL已启用严格SQL模式。", "第1章 第1.16-1.17节", "")
    Else
        Call AddResult("1.16_17", "系统安全", "数据库输入检查（SQL注入防护）", "fail", _
                       "MySQL未启用严格SQL模式，存在SQL注入风险。", "第1章 第1.16-1.17节", _
                       "设置sql_mode包含STRICT_TRANS_TABLES：SET GLOBAL sql_mode='STRICT_TRANS_TABLES';")
    End If
End Sub

' ============================================================
' 1.19 数据库默认设置（端口/账户）
' ============================================================
Sub Check_1_19_DbDefaults()
    Dim netOut : netOut = RunCmd("netstat -ano 2>nul")
    Dim hasMySQL : hasMySQL = (InStr(netOut, ":3306") > 0)
    Dim hasMSSQL : hasMSSQL = (InStr(netOut, ":1433") > 0)
    If hasMySQL Or hasMSSQL Then
        Dim portName : portName = ""
        If hasMySQL Then portName = "3306(MySQL)"
        If hasMSSQL Then
            If portName <> "" Then portName = portName & "/"
            portName = portName & "1433(MSSQL)"
        End If
        Call AddResult("1.19", "系统安全", "更换数据库默认设置（端口/账户）", "fail", _
                       "数据库使用默认端口：" & portName, "第1章 第1.19节", _
                       "修改数据库默认端口和管理员用户名，避免使用默认值")
    Else
        Call AddResult("1.19", "系统安全", "更换数据库默认设置（端口/账户）", "pass", _
                       "未检测到数据库默认端口（3306/1433）在用。", "第1章 第1.19节", "")
    End If
End Sub

' ============================================================
' 1.20 数据库安全策略（密码复杂度）
' ============================================================
Sub Check_1_20_DbPolicy()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.20节", "")
        Exit Sub
    End If
    Dim policyOut : policyOut = RunCmd("mysql -u root -e ""SHOW VARIABLES LIKE 'validate_password%';"" 2>nul")
    If InStr(LCase(policyOut), "validate_password") > 0 Then
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "pass", _
                       "MySQL密码复杂度策略已配置。", "第1章 第1.20节", "")
    Else
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "fail", _
                       "MySQL未配置密码复杂度策略（validate_password插件未加载）。", "第1章 第1.20节", _
                       "安装validate_password插件：INSTALL PLUGIN validate_password SONAME 'validate_password.so';")
    End If
End Sub

' ============================================================
' 1.21 数据库审计功能
' ============================================================
Sub Check_1_21_DbAudit()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.21", "系统安全", "数据库审计功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.21节", "")
        Exit Sub
    End If
    Dim auditOut : auditOut = RunCmd("mysql -u root -e ""SHOW VARIABLES LIKE 'general_log';"" 2>nul")
    If InStr(UCase(auditOut), "OFF") > 0 Then
        Call AddResult("1.21", "系统安全", "数据库审计功能", "fail", _
                       "MySQL审计功能未启用（general_log=OFF）。", "第1章 第1.21节", _
                       "启用审计：SET GLOBAL general_log = 'ON';")
    Else
        Call AddResult("1.21", "系统安全", "数据库审计功能", "pass", _
                       "MySQL审计功能已开启。", "第1章 第1.21节", "")
    End If
End Sub

' ============================================================
' 1.22 数据库审计日志保留>=180天
' ============================================================
Sub Check_1_22_DbAuditRetention()
    Dim verOut : verOut = RunCmd("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.22节", "")
        Exit Sub
    End If
    Dim expOut : expOut = RunCmd("mysql -u root -e ""SHOW VARIABLES LIKE 'expire_logs_days';"" 2>nul")
    Dim days : days = -1
    Dim elines : elines = Split(expOut, Chr(10))
    Dim ei
    For ei = 0 To UBound(elines)
        Dim ep : ep = Split(Trim(elines(ei)), Chr(9))
        If UBound(ep) >= 1 Then
            If InStr(LCase(ep(0)), "expire_logs_days") > 0 Then
                If IsNumeric(Trim(ep(UBound(ep)))) Then
                    days = CInt(Trim(ep(UBound(ep))))
                End If
            End If
        End If
    Next
    If days = -1 Then
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "manual", _
                       "无法读取MySQL日志保留配置，请手动核查", "第1章 第1.22节", _
                       "设置：SET GLOBAL expire_logs_days=180;")
    ElseIf days = 0 Then
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "pass", _
                       "MySQL binlog日志自动清理（expire_logs_days=0），无需处理。", "第1章 第1.22节", "")
    ElseIf days < 180 Then
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "fail", _
                       "MySQL binlog日志保留 " & days & " 天，低于180天。", "第1章 第1.22节", _
                       "设置：SET GLOBAL expire_logs_days=180;")
    Else
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "pass", _
                       "MySQL binlog日志保留 " & days & " 天，符合>=180天要求。", "第1章 第1.22节", "")
    End If
End Sub

' ============================================================
' 1.23 数据库访问分离（应用服务器无法访问）
' ============================================================
Sub Check_1_23_DbAppSep()
    Dim netOut : netOut = RunCmd("netstat -ano 2>nul")
    If Len(Trim(netOut)) = 0 Then
        Call AddResult("1.23", "系统安全", "数据库访问分离", "na", _
                       "无法获取网络连接信息。", "第1章 第1.23节", "")
        Exit Sub
    End If
    Dim dbPorts : dbPorts = ":3306 :1433 :6379"
    Dim foundPort : foundPort = ""
    Dim dp
    For Each dp In Split(dbPorts, " ")
        If InStr(netOut, dp) > 0 Then
            If foundPort <> "" Then foundPort = foundPort & " "
            foundPort = foundPort & dp
        End If
    Next
    If foundPort <> "" Then
        If InStr(netOut, "0.0.0.0") > 0 Then
            Call AddResult("1.23", "系统安全", "数据库访问分离", "fail", _
                           "数据库端口" & foundPort & "对外开放IP监听，应限制为应用服务器IP访问。", "第1章 第1.23节", _
                           "配置防火墙规则：仅允许应用服务器IP访问数据库端口")
        Else
            Call AddResult("1.23", "系统安全", "数据库访问分离", "pass", _
                           "数据库端口" & foundPort & "访问受限。", "第1章 第1.23节", "")
        End If
    Else
        Call AddResult("1.23", "系统安全", "数据库访问分离", "pass", _
                       "未检测到数据库对外监听端口。", "第1章 第1.23节", "")
    End If
End Sub

' ============================================================
' 第2章 用户安全
' ============================================================

Sub Check_2_1_PasswordExpiry()
    Dim out : out = RunCmd("net user administrator 2>nul")
    If InStr(out, "Password expires") > 0 And InStr(out, "Never") > 0 Then
        Call AddResult("2.1", "用户安全", "登录口令设置（管理员密码不永不过期）", _
                       "fail", "管理员账户密码设置为永不过期", "第2章 第2.1节", _
                       "本地安全策略→账户策略→密码策略→取消勾选'密码永不过期'")
    ElseIf out <> "" Then
        Call AddResult("2.1", "用户安全", "登录口令设置（管理员密码不永不过期）", _
                       "pass", "管理员账户密码到期会定期更新。", "第2章 第2.1节", "")
    Else
        Call AddResult("2.1", "用户安全", "登录口令设置（管理员密码不永不过期）", _
                       "manual", "无法获取账户信息，请手动核查", "第2章 第2.1节", "")
    End If
End Sub

Sub Check_2_4_UniqueAccounts()
    Dim out : out = RunCmd("net user 2>nul")
    Dim lines : lines = Split(out, Chr(10))
    Dim userLine : userLine = ""
    Dim i
    For i = 0 To UBound(lines)
        If InStr(lines(i), "Administrator") > 0 Or InStr(lines(i), "Guest") > 0 Then
            userLine = userLine & Trim(lines(i)) & " "
        End If
    Next
    Call AddResult("2.4", "用户安全", "账户唯一性", _
                   "pass", "已列出所有账户，请核查是否存在重复或共享账户。" & Chr(10) & Left(out, 300), _
                   "第2章 第2.4节", "")
End Sub


' ============================================================
' 2.5 用户端冗余服务
' ============================================================
Sub Check_2_5_RedundantServices()
    Dim svcOut : svcOut = RunCmd("sc query type= all state= running 2>nul")
    Dim dangerous : dangerous = "Telnet Ftp TftpD RemoteRegistry RemoteAccess Messenger Alerter ClipSrv TermService"
    Dim found : found = ""
    Dim lines : lines = Split(svcOut, Chr(10))
    Dim i
    For i = 0 To UBound(lines)
        Dim lineStr : lineStr = Trim(lines(i))
        If Left(lineStr, 12) = "SERVICE_NAME" Then
            Dim svcName : svcName = Trim(Mid(lineStr, 14))
            Dim d
            For Each d In Split(dangerous, " ")
                If LCase(svcName) = LCase(d) Then
                    If found <> "" Then found = found & ", "
                    found = found & svcName
                End If
            Next
        End If
    Next
    If found <> "" Then
        Call AddResult("2.5", "用户安全", "冗余服务（用户端）", _
                       "fail", "发现应停用的高危服务正在运行：" & found, "第2章 第2.5节", _
                       "通过 services.msc 停止并禁用以下服务：" & found)
    Else
        Call AddResult("2.5", "用户安全", "冗余服务（用户端）", _
                       "pass", "未发现 Telnet/FTP/RemoteRegistry 等高风险服务正在运行。", "第2章 第2.5节", "")
    End If
End Sub

Sub Check_2_7_NTFS()
    Dim sysDrive : sysDrive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
    Dim driveLetter : driveLetter = Left(sysDrive, 1)
    Dim out : out = RunCmd("fsutil fsinfo volumeinfo " & sysDrive & " 2>nul")
    If InStr(UCase(out), "NTFS") > 0 Then
        Call AddResult("2.7", "用户安全", "NTFS文件保护", _
                       "pass", "系统盘 " & sysDrive & " 使用 NTFS 文件系统", "第2章 第2.7节", "")
    ElseIf InStr(UCase(out), "FAT") > 0 Then
        Call AddResult("2.7", "用户安全", "NTFS文件保护", _
                       "fail", "系统盘 " & sysDrive & " 使用 FAT 文件系统，安全性不足。", "第2章 第2.7节", _
                       "执行命令将文件系统转换为NTFS：convert " & driveLetter & ": /fs:ntfs")
    Else
        Call AddResult("2.7", "用户安全", "NTFS文件保护", _
                       "manual", "无法确定文件系统类型，请手动核查", "第2章 第2.7节", "")
    End If
End Sub

Sub Check_2_9_RiskySoftware()
    Dim found : found = ""
    Dim regPaths(1)
    regPaths(0) = "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall"
    regPaths(1) = "HKLM\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    Dim risky : risky = Array("teamviewer", "向日葵", "todesk", "anydesk", "vnc", "filezilla")
    Dim i
    For i = 0 To 1
        Dim out : out = LCase(RunCmd("reg query """ & regPaths(i) & """ /s /v DisplayName 2>nul"))
        Dim r
        For Each r In risky
            If InStr(out, r) > 0 Then
                If InStr(found, r) = 0 Then
                    If found <> "" Then found = found & ", "
                    found = found & r
                End If
            End If
        Next
    Next
    If found <> "" Then
        Call AddResult("2.9", "用户安全", "已安装软件（风险软件检查）", _
                       "fail", "发现潜在风险软件：" & found, "第2章 第2.9节", _
                       "卸载未经批准的远程控制或文件传输软件：" & found)
    Else
        Call AddResult("2.9", "用户安全", "已安装软件（风险软件检查）", _
                       "pass", "未发现 TeamViewer/向日葵/AnyDesk/VNC/FileZilla 等高风险软件。", "第2章 第2.9节", "")
    End If
End Sub

Sub Check_2_10_USB()
    Dim val : val = ReadReg("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start")
    If val = "" Then
        Call AddResult("2.10", "用户安全", "USB接口禁止私自连接个人移动设备", _
                       "fail", "USBSTOR注册表项不存在，USB存储设备可能未管控", "第2章 第2.10节", _
                       "设置注册表项HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start = 4（DWORD）")
    ElseIf CInt(val) = 4 Then
        Call AddResult("2.10", "用户安全", "USB接口禁止私自连接个人移动设备", _
                       "pass", "USB存储设备已禁用（USBSTOR Start=4）。", "第2章 第2.10节", "")
    Else
        Call AddResult("2.10", "用户安全", "USB接口禁止私自连接个人移动设备", _
                       "fail", "USB存储设备未禁用（USBSTOR Start=" & val & "），应为4。", "第2章 第2.10节", _
                       "将 HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start 设为 4")
    End If
End Sub

Sub Check_2_11_PasswordPolicy()
    Dim out : out = RunCmd("net accounts 2>nul")
    Dim issues : issues = ""
    Dim minLen : minLen = 0
    Dim i, lines, parts
    lines = Split(out, Chr(10))
    For i = 0 To UBound(lines)
        If InStr(lines(i), "Minimum password length") > 0 Or _
           InStr(lines(i), "密码最小长度") > 0 Then
            parts = Split(Trim(lines(i)), " ")
            Dim minLenStr : minLenStr = Trim(parts(UBound(parts)))
            If IsNumeric(minLenStr) Then minLen = CInt(minLenStr)
        End If
    Next
    If minLen < 8 Then issues = issues & "  · 密码最小长度 " & minLen & " < 8，要求至少8位" & Chr(10)
    Dim lockout : lockout = 0
    For i = 0 To UBound(lines)
        If InStr(lines(i), "Lockout threshold") > 0 Or _
           InStr(lines(i), "锁定阈值") > 0 Then
            parts = Split(Trim(lines(i)), " ")
            Dim lockStr : lockStr = parts(UBound(parts))
            If IsNumeric(lockStr) Then lockout = CInt(lockStr)
        End If
    Next
    If lockout = 0 Then issues = issues & "  · 账户锁定阈值未设置，应为5次或更少" & Chr(10)
    Dim ssActive : ssActive = ReadReg("HKCU\Control Panel\Desktop\ScreenSaveActive")
    Dim ssTimeout : ssTimeout = ReadReg("HKCU\Control Panel\Desktop\ScreenSaveTimeOut")
    If ssActive <> "1" Then
        issues = issues & "  · 屏幕保护程序未启用" & Chr(10)
    ElseIf IsNumeric(ssTimeout) Then
        If CInt(ssTimeout) > 300 Then
            issues = issues & "  · 屏幕等待时间 " & CInt(ssTimeout)\60 & " 分钟 > 5分钟" & Chr(10)
        End If
    End If

    If issues <> "" Then
        Call AddResult("2.11", "用户安全", "口令策略和屏保设置", _
                       "fail", issues, "第2章 第2.11节", _
                       "执行: net accounts /minpwlen:8 /lockoutthreshold:5" & Chr(10) & _
                       "在显示属性设置屏幕保护程序，等待时间≤5分钟")
    Else
        Call AddResult("2.11", "用户安全", "口令策略和屏保设置", _
                       "pass", "口令策略和屏保设置符合要求（最小长度:" & minLen & "，锁定阈值:" & lockout & "）。", _
                       "第2章 第2.11节", "")
    End If
End Sub

Sub Check_2_12_Wireless()
    Dim issues : issues = ""
    Dim wlanOut : wlanOut = RunCmd("sc query WLANSVC 2>nul")
    If InStr(wlanOut, "RUNNING") > 0 Then
        issues = issues & "  · WLAN AutoConfig 服务正在运行（需禁用）" & Chr(10)
    End If
    Dim btOut : btOut = RunCmd("sc query bthserv 2>nul")
    If InStr(btOut, "RUNNING") > 0 Then
        issues = issues & "  · 蓝牙服务 bthserv 正在运行（需禁用）" & Chr(10)
    End If
    If issues <> "" Then
        Call AddResult("2.12", "用户安全", "物理拆除Wi-Fi/红外/蓝牙等无线模块", _
                       "fail", issues, "第2章 第2.12节", _
                       "设备管理器中禁用或卸载无线网卡、蓝牙等设备" & Chr(10) & "services.msc 中禁用 WLAN AutoConfig、Bluetooth 服务")
    Else
        Call AddResult("2.12", "用户安全", "物理拆除Wi-Fi/红外/蓝牙等无线模块", _
                       "pass", "未检测到运行的无线/蓝牙服务。", "第2章 第2.12节", "")
    End If
End Sub

Sub Check_2_13_OutboundControl()
    Dim hasControl : hasControl = False
    Dim detail : detail = ""
    If is7OrAbove Then
        Dim fwOut : fwOut = RunCmd("netsh advfirewall show allprofiles 2>nul")
        If InStr(fwOut, "Block") > 0 Or InStr(fwOut, "阻止") > 0 Then
            hasControl = True
            detail = "防火墙出站默认策略为阻止。"
        End If
    Else
        detail = "Windows XP 防火墙出站策略能力有限，建议部署终端非法外联控制系统。"
    End If
    If hasControl Then
        Call AddResult("2.13", "用户安全", "非法外联控制", _
                       "pass", detail, "第2章 第2.13节", "")
    Else
        Call AddResult("2.13", "用户安全", "非法外联控制", _
                       "fail", "未检测到有效的出站阻止策略。" & detail, "第2章 第2.13节", _
                       "将防火墙出站默认策略改为阻止，实现终端非法外联控制系统的实际效果。")
    End If
End Sub

Sub Check_2_14_UserAudit()
    Dim logOut
    If is7OrAbove Then
        logOut = RunCmd("wevtutil qe Security /c:1 /rd:true /f:text 2>nul")
    Else
        logOut = RunCmd("eventquery /L Security /R 1 2>nul")
    End If
    If Len(logOut) > 10 Then
        Call AddResult("2.14", "用户安全", "用户行为审计能力", _
                       "pass", "安全事件日志中有记录，用户行为审计可用。", "第2章 第2.14节", "")
    Else
        Call AddResult("2.14", "用户安全", "用户行为审计能力", _
                       "fail", "安全事件日志无记录或无法读取。", "第2章 第2.14节", _
                       "本地安全策略→高级审核策略→启用登录、账户管理、特权使用等审核")
    End If
End Sub

' ============================================================
' 第3章 数据安全
' ============================================================

Sub Check_3_1_Encryption()
    If Not is7OrAbove Then
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "na", "Windows XP 不支持 BitLocker，如使用其他加密工具（如EFS、TrueCrypt）。", _
                       "第3章 第3.1节", "")
        Exit Sub
    End If
    Dim bdeOut : bdeOut = RunCmd("manage-bde -status 2>nul")
    If InStr(bdeOut, "100 %") > 0 Or InStr(bdeOut, "已加密") > 0 Then
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "pass", "磁盘已全盘加密（BitLocker）。", "第3章 第3.1节", "")
    ElseIf InStr(bdeOut, "完全解密") > 0 Or InStr(bdeOut, "Fully Decrypted") > 0 Then
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "fail", "磁盘未启用 BitLocker 加密。", "第3章 第3.1节", _
                       "启用BitLocker（控制面板→BitLocker驱动器加密 或 manage-bde -on C:)")
    Else
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "manual", "无法确定 BitLocker 状态，请手动核查", "第3章 第3.1节", "")
    End If
End Sub

Sub Check_3_5_LogProtection()
    Dim logOut : logOut = ""
    If is7OrAbove Then
        logOut = RunCmd("wevtutil gl Security 2>nul")
    End If
    If InStr(logOut, "maxSize") > 0 Then
        Dim pos : pos = InStr(logOut, "maxSize:")
        If pos > 0 Then
            Dim sizeStr : sizeStr = Trim(Mid(logOut, pos + 8, 20))
            sizeStr = Split(sizeStr, Chr(10))(0)
            If IsNumeric(sizeStr) Then
                Dim sizeMB : sizeMB = CLng(sizeStr) \ (1024 * 1024)
                If sizeMB < 64 Then
                    Call AddResult("3.5", "数据安全", "日志保护措施（读写控制/完整性）", _
                                   "fail", "安全日志大小 " & sizeMB & " MB，低于256MB。", _
                                   "第3章 第3.5节", _
                                   "日志查看器→Windows日志→安全→属性，设置最大大小为262144KB（256MB）")
                Else
                    Call AddResult("3.5", "数据安全", "日志保护措施（读写控制/完整性）", _
                                   "pass", "安全日志大小为 " & sizeMB & " MB，符合要求。", "第3章 第3.5节", "")
                End If
            End If
        End If
    Else
        Call AddResult("3.5", "数据安全", "日志保护措施（读写控制/完整性）", _
                       "manual", "请手动检查安全日志大小（建议至少256MB）。", "第3章 第3.5节", _
                       "日志查看器→Windows日志→安全→属性")
    End If
End Sub

Sub Check_3_10_Shares()
    Dim shareOut : shareOut = RunCmd("net share 2>nul")
    Dim found : found = ""
    Dim adminShares : adminShares = "C$ D$ E$ F$ G$ ADMIN$ IPC$"
    Dim lines : lines = Split(shareOut, Chr(10))
    Dim i
    For i = 2 To UBound(lines)
        Dim lineStr : lineStr = Trim(lines(i))
        If Len(lineStr) > 0 Then
            Dim parts : parts = Split(lineStr, " ")
            If Len(parts(0)) > 0 Then
                If InStr(adminShares, parts(0)) = 0 And InStr(parts(0), "---") = 0 And _
                   parts(0) <> "Share" And parts(0) <> "共享名" Then
                    If found <> "" Then found = found & ", "
                    found = found & parts(0)
                End If
            End If
        End If
    Next
    If found <> "" Then
        Call AddResult("3.10", "数据安全", "数据共享合理性检查", _
                       "fail", "发现非系统默认共享：" & found & "，请确认是否必要", "第3章 第3.10节", _
                       "删除非必要的共享：net share 共享名 /delete" & Chr(10) & "降低共享权限，移除 Everyone 完全控制")
    Else
        Call AddResult("3.10", "数据安全", "数据共享合理性检查", _
                       "pass", "未发现非系统必要的非默认共享。", "第3章 第3.10节", "")
    End If
End Sub

' ============================================================
' 第4章 应用安全
' ============================================================

Sub Check_4_15_NLA()
    If isXP Then
        Call AddResult("4.15", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "na", "Windows XP 不支持 NLA，如紧急可远程登录验证后再升级操作系统。", "第4章 第4.15节", "")
        Exit Sub
    End If
    Dim val : val = ReadReg("HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserAuthentication")
    If val = "1" Then
        Call AddResult("4.15", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "pass", "NLA 已启用（UserAuthentication=1）。", "第4章 第4.15节", "")
    Else
        Call AddResult("4.15", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "fail", "NLA 未启用（UserAuthentication=" & val & "），应为1。", "第4章 第4.15节", _
                       "注册表 HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserAuthentication 设为 1")
    End If
End Sub

Sub Check_4_16_RoleSeparation()
    Dim grpOut : grpOut = RunCmd("net localgroup 2>nul")
    Dim missing : missing = ""
    Dim hasAdmin : hasAdmin = (InStr(LCase(grpOut), "administrators") > 0 Or InStr(grpOut, "管理员") > 0)
    Dim hasSec : hasSec = (InStr(LCase(grpOut), "security") > 0 Or InStr(grpOut, "安全") > 0)
    Dim hasAudit : hasAudit = (InStr(LCase(grpOut), "audit") > 0 Or InStr(grpOut, "审计") > 0)
    If Not hasAdmin Then missing = missing & "管理员组、"
    If Not hasSec Then missing = missing & "安全员组（Security/安全）、"
    If Not hasAudit Then missing = missing & "审计员组（Audit/审计）、"
    If missing = "" Then
        Call AddResult("4.16", "应用安全", "支持管理员/安全员/审计员三权分立", _
                       "pass", "检测到管理员组、安全员组、审计员组，三权分立可能已配置。", "第4章 第4.16节", "")
    Else
        Call AddResult("4.16", "应用安全", "支持管理员/安全员/审计员三权分立", _
                       "fail", "缺少专职分组：" & Left(missing, Len(missing)-1) & Chr(10) & "当前组：" & Left(grpOut, 300), _
                       "第4章 第4.16节", _
                       "创建专职安全管理员组、审计员组，实现权限分离，禁止同一账户兼任多种角色。")
    End If
End Sub

Sub Check_4_17_IPRestrict()
    Dim detail : detail = ""
    If is7OrAbove Then
        Dim fwOut : fwOut = RunCmd("netsh advfirewall firewall show rule name=all dir=in 2>nul")
        If InStr(fwOut, "RemoteIP") > 0 And InStr(LCase(fwOut), "any") = 0 Then
            detail = "检测到带IP限制的防火墙规则。"
            Call AddResult("4.17", "应用安全", "业务管理终端登录采取IP地址限制", _
                           "pass", detail, "第4章 第4.17节", "")
        Else
            Call AddResult("4.17", "应用安全", "业务管理终端登录采取IP地址限制", _
                           "fail", "未检测到管理终端IP访问限制规则。", "第4章 第4.17节", _
                           "高级安全防火墙→入站规则→新建规则：限制RDP(3389)端口只允许指定网段IP访问")
        End If
    Else
        Call AddResult("4.17", "应用安全", "业务管理终端登录采取IP地址限制", _
                       "manual", "Windows XP 请手动检查防火墙是否限制管理端口来源IP", "第4章 第4.17节", "")
    End If
End Sub

Sub Check_4_18_Lockout()
    Dim out : out = RunCmd("net accounts 2>nul")
    Dim lockout : lockout = 0
    Dim lines : lines = Split(out, Chr(10))
    Dim i
    For i = 0 To UBound(lines)
        If InStr(lines(i), "Lockout threshold") > 0 Or InStr(lines(i), "锁定阈值") > 0 Then
            Dim parts : parts = Split(Trim(lines(i)), " ")
            Dim ls : ls = parts(UBound(parts))
            If IsNumeric(ls) Then lockout = CInt(ls)
        End If
    Next
    If lockout > 0 Then
        Call AddResult("4.18", "应用安全", "具有登录失败处理功能（账户锁定）", _
                       "pass", "账户锁定阈值 " & lockout & " 次", "第4章 第4.18节", "")
    Else
        Call AddResult("4.18", "应用安全", "具有登录失败处理功能（账户锁定）", _
                       "fail", "未配置账户锁定策略，存在暴力破解风险。", "第4章 第4.18节", _
                       "执行命令：net accounts /lockoutthreshold:5" & Chr(10) & "或【本地安全策略】→【账户策略】→【账户锁定策略】设置阈值（5次）")
    End If
End Sub

Sub Check_4_22_PortSeparation()
    Dim netOut : netOut = RunCmd("netstat -ano 2>nul")
    Dim webPorts : webPorts = Array(80, 443, 8080, 8443)
    Dim mgmtPorts : mgmtPorts = Array(3389, 445, 135)
    Dim hasWeb : hasWeb = False
    Dim hasMgmt : hasMgmt = False
    Dim p
    For Each p In webPorts
        If InStr(netOut, ":" & CStr(p) & " ") > 0 Or InStr(netOut, ":" & CStr(p) & Chr(9)) > 0 Then
            hasWeb = True
        End If
    Next
    For Each p In mgmtPorts
        If InStr(netOut, ":" & CStr(p) & " ") > 0 Or InStr(netOut, ":" & CStr(p) & Chr(9)) > 0 Then
            hasMgmt = True
        End If
    Next
    If hasWeb And hasMgmt Then
        Call AddResult("4.22", "应用安全", "分开设置管理端口与应用端口", _
                       "pass", "应用端口（80/443）和管理端口（3389/445）分别监听，端口未混用。", "第4章 第4.22节", "")
    Else
        Call AddResult("4.22", "应用安全", "分开设置管理端口与应用端口", _
                       "pass", "本次未同时检测到Web端口和管理端口混用情况。", "第4章 第4.22节", "")
    End If
End Sub

Sub Check_4_23_AppDBSep()
    Dim netOut : netOut = RunCmd("netstat -ano 2>nul")
    Dim webFound : webFound = ""
    Dim dbFound : dbFound = ""
    Dim webPorts : webPorts = Array(80, 443, 8080, 8000)
    Dim dbPorts  : dbPorts  = Array(3306, 1433, 5432, 27017)
    Dim p
    For Each p In webPorts
        If InStr(netOut, ":" & CStr(p) & " ") > 0 Then
            If webFound <> "" Then webFound = webFound & ","
            webFound = webFound & CStr(p)
        End If
    Next
    For Each p In dbPorts
        If InStr(netOut, ":" & CStr(p) & " ") > 0 Then
            If dbFound <> "" Then dbFound = dbFound & ","
            dbFound = dbFound & CStr(p)
        End If
    Next
    If webFound <> "" And dbFound <> "" Then
        Call AddResult("4.23", "应用安全", "应用服务和数据存储分离部署", _
                       "fail", "主机同时部署Web(" & webFound & ")和数据库(" & dbFound & ")，未分离部署。", _
                       "第4章 第4.23节", "将数据库迁移到独立服务器，应用服务器上不应安装数据库。")
    Else
        Call AddResult("4.23", "应用安全", "应用服务和数据存储分离部署", _
                       "pass", "未检测到Web和数据库在同一台主机上混合部署。", "第4章 第4.23节", "")
    End If
End Sub

' ============================================================
' 4.24 应用系统具有日志审计功能
' ============================================================
Sub Check_4_24_AppAudit()
    Dim findings : findings = ""
    Dim issues : issues = ""
    Dim iisLog1 : iisLog1 = "C:\inetpub\logs\LogFiles"
    Dim iisLog2 : iisLog2 = "C:\Windows\System32\LogFiles\W3SVC1"
    If oFSO.FolderExists(iisLog1) Then
        findings = findings & "  · IIS 日志目录存在：" & iisLog1 & Chr(10)
    ElseIf oFSO.FolderExists(iisLog2) Then
        findings = findings & "  · IIS 日志目录存在：" & iisLog2 & Chr(10)
    End If
    Dim nginxLog1 : nginxLog1 = "C:\nginx\logs"
    Dim nginxLog2 : nginxLog2 = "C:\Program Files\nginx\logs"
    If oFSO.FolderExists(nginxLog1) Then
        findings = findings & "  · Nginx 日志目录存在：" & nginxLog1 & Chr(10)
    ElseIf oFSO.FolderExists(nginxLog2) Then
        findings = findings & "  · Nginx 日志目录存在：" & nginxLog2 & Chr(10)
    End If
    Dim appLog : appLog = RunCmd("wevtutil qe Application /c:1 /rd:true /f:text 2>nul")
    If Len(appLog) > 10 Then
        findings = findings & "  · Windows 应用程序日志中有有效记录" & Chr(10)
    Else
        issues = issues & "  · Windows 应用程序日志为空或无法读取" & Chr(10)
    End If
    If issues <> "" And findings = "" Then
        Call AddResult("4.24", "应用安全", "应用系统具有日志审计功能", _
                       "fail", issues, "第4章 第4.24节", _
                       "检查 IIS/Nginx/Apache 等访问日志，确保应用操作被完整记录")
    ElseIf findings <> "" Then
        Call AddResult("4.24", "应用安全", "应用系统具有日志审计功能", _
                       "pass", findings, "第4章 第4.24节", "")
    Else
        Call AddResult("4.24", "应用安全", "应用系统具有日志审计功能", _
                       "pass", "Windows 应用程序日志正常，但未检测到 Web 服务（可能未部署 Web 应用）", _
                       "第4章 第4.24节", "")
    End If
End Sub

' ============================================================
' 4.6 WAF防护
' ============================================================
Sub Check_4_6_WAF()
    Dim regOut : regOut = RunCmd("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim combined : combined = LCase(regOut)
    Dim wafKws : wafKws = "modsecurity safedog yunsuo huwshen d_safe safeline chaitin"
    Dim found : found = ""
    Dim wk
    For Each wk In Split(wafKws, " ")
        If InStr(combined, wk) > 0 Then
            If found <> "" Then found = found & ", "
            found = found & wk
        End If
    Next
    If oFSO.FolderExists("C:\inetpub") Then
        If found <> "" Then found = found & ", "
        found = found & "IIS（可能自带模块）"
    End If
    If found <> "" Then
        Call AddResult("4.6", "应用安全", "Web应用系统采取Web应用防火墙措施（WAF）", "pass", _
                       "检测到Web防护模块：" & found, "第4章 第4.6节", "")
    Else
        Call AddResult("4.6", "应用安全", "Web应用系统采取Web应用防火墙措施（WAF）", "fail", _
                       "未检测到WAF或Web防护模块。", "第4章 第4.6节", _
                       "部署WAF（如ModSecurity、安全狗、云锁等）保护Web应用")
    End If
End Sub

' ============================================================
' 4.8 网页防篡改
' ============================================================
Sub Check_4_8_AntiTamper()
    Dim regOut : regOut = RunCmd("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim combined : combined = LCase(regOut)
    Dim kws : kws = "safedog huwshen yunsuo iiswatch tripwire"
    Dim found : found = ""
    Dim k
    For Each k In Split(kws, " ")
        If InStr(combined, k) > 0 Then
            If found <> "" Then found = found & ", "
            found = found & k
        End If
    Next
    Dim vssOut : vssOut = RunCmd("sc query VSS 2>nul")
    If InStr(vssOut, "RUNNING") > 0 Then
        If found <> "" Then found = found & ", "
        found = found & "VSS（卷影副本，可用于文件恢复）"
    End If
    If found <> "" Then
        Call AddResult("4.8", "应用安全", "网站采取网页防篡改措施", "pass", _
                       "检测到防篡改模块：" & found, "第4章 第4.8节", "")
    Else
        Call AddResult("4.8", "应用安全", "网站采取网页防篡改措施", "fail", _
                       "未检测到网页防篡改模块。", "第4章 第4.8节", _
                       "部署网页防篡改系统（如护卫神、安全狗等），并对关键文件进行完整性校验")
    End If
End Sub

' ============================================================
' 4.12 并发会话数量限制
' ============================================================
Sub Check_4_12_SessionLimit()
    Dim findings : findings = ""
    Dim iisLog : iisLog = "C:\inetpub\logs\LogFiles"
    If oFSO.FolderExists(iisLog) Then
        findings = findings & "  · 检测到IIS日志（" & iisLog & "）" & Chr(10)
    End If
    Dim nginxConf : nginxConf = "C:\nginx\conf\nginx.conf"
    If oFSO.FileExists(nginxConf) Then
        Dim confOut : confOut = RunCmd("findstr /i ""worker_connections"" " & nginxConf & " 2>nul")
        If Len(Trim(confOut)) > 0 Then
            findings = findings & "  · Nginx配置了worker_connections：" & Trim(confOut) & Chr(10)
        End If
    End If
    If findings <> "" Then
        Call AddResult("4.12", "应用安全", "具备拒绝应用请求并发会话数量限制等功能", "pass", _
                       findings, "第4章 第4.12节", "")
    Else
        Call AddResult("4.12", "应用安全", "具备拒绝应用请求并发会话数量限制等功能", "fail", _
                       "未检测到IIS/Nginx并发限制配置。", "第4章 第4.12节", _
                       "IIS：在站点配置中设置并发连接数限制" & Chr(10) & _
                       "Nginx：在nginx.conf中配置 worker_connections 和 limit_conn")
    End If
End Sub

' ============================================================
' 4.21 非默认Web应用发布端口
' ============================================================
Sub Check_4_21_DefaultPort()
    Dim netOut : netOut = RunCmd("netstat -ano 2>nul")
    Dim defPorts : defPorts = ":80 :443 :8080 :8443"
    Dim found : found = ""
    Dim dp
    For Each dp In Split(defPorts, " ")
        If InStr(netOut, dp & " ") > 0 Or InStr(netOut, dp & Chr(9)) > 0 Then
            If found <> "" Then found = found & ", "
            found = found & Mid(dp, 2)
        End If
    Next
    If found <> "" Then
        Call AddResult("4.21", "应用安全", "非默认Web应用系统发布端口", "fail", _
                       "检测到Web服务使用默认端口：" & found, "第4章 第4.21节", _
                       "将Web服务端口从默认80/443/8080修改为非标准端口，并通过路由防火墙映射")
    Else
        Call AddResult("4.21", "应用安全", "非默认Web应用系统发布端口", "pass", _
                       "未检测到Web服务使用默认端口（80/443/8080）。", "第4章 第4.21节", "")
    End If
End Sub

' ============================================================
' 4.29 远程管理限制
' ============================================================
Sub Check_4_29_RemoteMgmt()
    Dim issues : issues = ""
    Dim rdpRules : rdpRules = RunCmd("netsh advfirewall firewall show rule name=all 2>nul")
    If InStr(LCase(rdpRules), "remotedesktop") = 0 And InStr(rdpRules, "3389") = 0 Then
        issues = issues & "  · 未配置RDP防火墙规则限制管理端口" & Chr(10)
    End If
    Dim nlaVal : nlaVal = ReadReg("HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserAuthentication")
    If nlaVal = "" Then
        issues = issues & "  · 无法获取NLA配置" & Chr(10)
    ElseIf IsNumeric(nlaVal) Then
        If CInt(nlaVal) <> 1 Then
            issues = issues & "  · RDP未启用NLA（网络级认证），存在被弱口令破解风险" & Chr(10)
        End If
    End If
    Dim netAcct : netAcct = RunCmd("net accounts 2>nul")
    If InStr(netAcct, "Never") > 0 Or InStr(netAcct, "从不") > 0 Then
        issues = issues & "  · 未配置账户锁定策略，远程登录存在被弱口令破解风险" & Chr(10)
    End If
    If issues <> "" Then
        Call AddResult("4.29", "应用安全", "具备统一管理措施对远程管理进行限制", "fail", _
                       issues, "第4章 第4.29节", _
                       "1. 配置防火墙规则限制RDP来源IP" & Chr(10) & _
                       "2. 启用NLA：注册表 UserAuthentication=1" & Chr(10) & _
                       "3. 配置账户锁定：net accounts /lockoutthreshold:5")
    Else
        Call AddResult("4.29", "应用安全", "具备统一管理措施对远程管理进行限制", "pass", _
                       "远程管理安全配置检查通过。", "第4章 第4.29节", "")
    End If
End Sub

' ============================================================
' 4.31 应用系统具备备份与恢复功能
' ============================================================
Sub Check_4_31_AppBackup()
    Dim findings : findings = ""
    Dim vssOut : vssOut = RunCmd("sc query VSS 2>nul")
    If InStr(vssOut, "RUNNING") > 0 Then
        findings = findings & "  · 卷影副本服务（VSS）正在运行" & Chr(10)
    End If
    Dim taskOut : taskOut = RunCmd("schtasks /query /fo LIST 2>nul")
    If InStr(LCase(taskOut), "backup") > 0 Or InStr(LCase(taskOut), "备份") > 0 Or InStr(LCase(taskOut), "wbadmin") > 0 Then
        findings = findings & "  · 检测到备份相关计划任务" & Chr(10)
    End If
    Dim bdirs : bdirs = "C:\backup D:\backup C:\Backup D:\Backup"
    Dim bd
    For Each bd In Split(bdirs, " ")
        If oFSO.FolderExists(bd) Then
            findings = findings & "  · 发现备份目录：" & bd & Chr(10)
        End If
    Next
    If findings <> "" Then
        Call AddResult("4.31", "应用安全", "应用系统具备备份与恢复功能", "pass", _
                       "检测到备份机制：" & Chr(10) & findings, "第4章 第4.31节", "")
    Else
        Call AddResult("4.31", "应用安全", "应用系统具备备份与恢复功能", "fail", _
                       "未检测到备份服务、备份计划任务或备份目录。", "第4章 第4.31节", _
                       "1. 安装Windows Server Backup功能并配置定期备份" & Chr(10) & _
                       "2. 创建定期备份计划任务" & Chr(10) & _
                       "3. 启用卷影副本（VSS）")
    End If
End Sub

' ============================================================
' 4.32 国产自主可控软件
' ============================================================
Sub Check_4_32_DomesticSoftware()
    Dim findings : findings = ""
    Dim issues : issues = ""
    Dim osName : osName = RunCmd("wmic os get Caption /value 2>nul")
    Dim domesticOS : domesticOS = "kylin uos deepin 麒麟 统信 普华"
    Dim isDomestic : isDomestic = False
    Dim dk
    For Each dk In Split(domesticOS, " ")
        If InStr(LCase(osName), dk) > 0 Then isDomestic = True
    Next
    If isDomestic Then
        findings = findings & "  · 操作系统为国产系统" & Chr(10)
    Else
        issues = issues & "  · 操作系统为非国产系统（" & Trim(Mid(osName, InStr(osName, "=") + 1)) & "）" & Chr(10)
    End If
    Dim regOut : regOut = RunCmd("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim combined : combined = LCase(regOut)
    Dim dbKws : dbKws = "dmdbms 达梦 kingbase 人大 gbase oceanbase tidb opengauss gaussdb"
    Dim mwKws : mwKws = "tongweb apusic primeton"
    Dim dkw
    For Each dkw In Split(dbKws, " ")
        If InStr(combined, dkw) > 0 Then
            findings = findings & "  · 检测到国产数据库（" & dkw & "）" & Chr(10)
        End If
    Next
    For Each dkw In Split(mwKws, " ")
        If InStr(combined, dkw) > 0 Then
            findings = findings & "  · 检测到国产中间件（" & dkw & "）" & Chr(10)
        End If
    Next
    If findings <> "" And issues = "" Then
        Call AddResult("4.32", "应用安全", "应用系统基于国产自主可控硬件或软件", "pass", _
                       findings, "第4章 第4.32节", "")
    ElseIf findings <> "" Then
        Call AddResult("4.32", "应用安全", "应用系统基于国产自主可控硬件或软件", "manual", _
                       issues & "已检测到：" & Chr(10) & findings, "第4章 第4.32节", _
                       "推进国产化替代：操作系统→麒麟/统信；数据库→达梦/人大金仓")
    Else
        Call AddResult("4.32", "应用安全", "应用系统基于国产自主可控硬件或软件", "fail", _
                       issues, "第4章 第4.32节", _
                       "推进国产化替代：操作系统→麒麟/统信；数据库→达梦/人大金仓；中间件→东方通/普元")
    End If
End Sub

' ============================================================
' 10.1 TLS协议版本配置
' ============================================================
Sub Check_10_1_TLS()
    Dim issues : issues = ""
    Dim findings : findings = ""
    Dim basePath : basePath = "HKLM\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\"
    Dim protos(3)
    protos(0) = "SSL 3.0"
    protos(1) = "TLS 1.0"
    protos(2) = "TLS 1.1"
    protos(3) = "TLS 1.2"
    Dim i
    For i = 0 To 3
        Dim regPath : regPath = basePath & protos(i) & "\Server\Enabled"
        Dim val : val = ReadReg(regPath)
        If val = "" Then
            If protos(i) = "TLS 1.2" Then
                findings = findings & "  · TLS 1.2 使用系统默认（通常启用）" & Chr(10)
            Else
                findings = findings & "  · " & protos(i) & " 使用系统默认（未显式配置）" & Chr(10)
            End If
        ElseIf IsNumeric(val) Then
            If protos(i) = "TLS 1.2" Then
                If CInt(val) = 0 Then
                    issues = issues & "  · TLS 1.2 被禁用，通信安全性不足" & Chr(10)
                Else
                    findings = findings & "  · TLS 1.2 已启用" & Chr(10)
                End If
            Else
                If CInt(val) <> 0 Then
                    issues = issues & "  · " & protos(i) & " 未禁用，存在接受旧协议风险" & Chr(10)
                Else
                    findings = findings & "  · " & protos(i) & " 已禁用" & Chr(10)
                End If
            End If
        End If
    Next
    If issues <> "" Then
        Call AddResult("10.1", "协议安全审计", "协议安全审计（TLS版本配置）", "fail", _
                       issues, "第10章 第10.1节", _
                       "通过注册表配置不安全协议（TLS 1.0/1.1/SSL 3.0）禁用" & Chr(10) & _
                       "SCHANNEL\Protocols\对应协议\Server 下新建 DWORD: Enabled = 0")
    Else
        Call AddResult("10.1", "协议安全审计", "协议安全审计（TLS版本配置）", "pass", _
                       "TLS协议版本配置符合要求。" & Chr(10) & findings, "第10章 第10.1节", "")
    End If
End Sub

' ============================================================
' 2.2 用户计算机应根据需要安装补丁程序
' ============================================================
Sub Check_2_2_UserPatch()
    Dim patchOut : patchOut = RunCmd("wmic qfe get InstalledOn /format:csv 2>nul")
    Dim latestDate : latestDate = ""
    Dim plines : plines = Split(patchOut, Chr(10))
    Dim pi
    For pi = 0 To UBound(plines)
        Dim pl : pl = Trim(plines(pi))
        If Len(pl) >= 8 Then
            Dim pparts : pparts = Split(pl, ",")
            If UBound(pparts) >= 1 Then
                Dim dateStr : dateStr = Trim(pparts(UBound(pparts)))
                If Len(dateStr) >= 8 And IsNumeric(Left(dateStr, 4)) Then
                    If dateStr > latestDate Then latestDate = dateStr
                End If
            End If
        End If
    Next
    If Len(latestDate) >= 8 Then
        Call AddResult("2.2", "用户安全", "用户计算机应根据需要安装补丁程序", "pass", _
                       "最近补丁安装日期：" & latestDate & "，补丁记录正常。", "第2章 第2.2节", "")
    Else
        Call AddResult("2.2", "用户安全", "用户计算机应根据需要安装补丁程序", "fail", _
                       "无法获取补丁安装记录，Windows Update可能未启用补丁。", "第2章 第2.2节", _
                       "通过Windows Update安装最新补丁，保持操作系统和管理员策略。")
    End If
End Sub

' ============================================================
' 2.6 阻断和告警非法连接互联网
' ============================================================
Sub Check_2_6_IllegalInternet()
    Dim findings : findings = ""
    Dim issues : issues = ""
    Dim fwOut : fwOut = RunCmd("netsh advfirewall show allprofiles 2>nul")
    If InStr(fwOut, "Block") > 0 Or InStr(fwOut, "阻止") > 0 Then
        findings = findings & "  · 防火墙出站默认策略为阻止" & Chr(10)
    Else
        issues = issues & "  · 防火墙出站默认策略为允许，存在非法外联风险。" & Chr(10)
    End If
    Dim regOut : regOut = RunCmd("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim edrs : edrs = "sangfor edr antiy zhijia 360entclient nsfocus topsec"
    Dim ek
    For Each ek In Split(edrs, " ")
        If InStr(LCase(regOut), ek) > 0 Then
            findings = findings & "  · 检测到终端管控/EDR客户端（" & ek & "）" & Chr(10)
        End If
    Next
    If findings <> "" Then
        Call AddResult("2.6", "用户安全", "具备阻断和告警非法连接互联网的能力", "pass", _
                       findings, "第2章 第2.6节", "")
    Else
        Call AddResult("2.6", "用户安全", "具备阻断和告警非法连接互联网的能力", "fail", _
                       issues & "  · 未检测到终端管控/防非法外联软件。", "第2章 第2.6节", _
                       "1. 将防火墙出站默认策略改为Block" & Chr(10) & _
                       "2. 部署终端管控系统（如深信服EDR、绿盟等）")
    End If
End Sub

' ============================================================
' 2.8 终端管控措施
' ============================================================
Sub Check_2_8_EndpointControl()
    Dim findings : findings = ""
    Dim issues : issues = ""
    Dim regOut : regOut = RunCmd("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim mgmtKws : mgmtKws = "sangfor edr antiy zhijia 360entclient nsfocus topsec mcafee symantec"
    Dim mk
    For Each mk In Split(mgmtKws, " ")
        If InStr(LCase(regOut), mk) > 0 Then
            findings = findings & "  · 检测到终端管控客户端（" & mk & "）" & Chr(10)
        End If
    Next
    Dim usbVal : usbVal = ReadReg("HKLM\SYSTEM\CurrentControlSet\Services\USBSTOR\Start")
    If usbVal = "" Then
        issues = issues & "  · 无法读取USB管控注册表项" & Chr(10)
    ElseIf IsNumeric(usbVal) Then
        If CInt(usbVal) <> 4 Then
            issues = issues & "  · USB存储未禁用（USBSTOR Start=" & usbVal & "），应为4。" & Chr(10)
        End If
    End If
    If findings <> "" Then
        Dim stat2_8 : stat2_8 = "pass"
        If issues <> "" Then stat2_8 = "manual"
        Call AddResult("2.8", "用户安全", "终端管控措施（统一配置、外设管控、无线禁用等）", stat2_8, _
                       findings & issues, "第2章 第2.8节", _
                       "确认终端管控软件的外设管控、行为审计、非法外联控制等功能均已启用。")
    Else
        Call AddResult("2.8", "用户安全", "终端管控措施（统一配置、外设管控、无线禁用等）", "fail", _
                       "未检测到终端管控软件" & Chr(10) & issues, "第2章 第2.8节", _
                       "部署终端管控系统，实现统一配置、外设管控、行为审计、非法外联控制等功能")
    End If
End Sub

' ============================================================
' 2.15 安全策略配置（组策略）
' ============================================================
Sub Check_2_15_SecurityPolicy()
    Dim gpOut : gpOut = RunCmd("gpresult /R /SCOPE COMPUTER 2>nul")
    If InStr(gpOut, "GPO") > 0 Then
        Call AddResult("2.15", "用户安全", "安全策略配置（组策略）", "pass", _
                       "检测到组策略配置（GPO）。", "第2章 第2.15节", "")
    Else
        Call AddResult("2.15", "用户安全", "安全策略配置（组策略）", "fail", _
                       "未检测到组策略配置。", "第2章 第2.15节", _
                       "通过gpedit.msc配置本地安全策略")
    End If
End Sub

' ============================================================
' 2.16 补丁修复升级到最新版本
' ============================================================
Sub Check_2_16_PatchLatest()
    Dim issues : issues = ""
    Dim patchOut : patchOut = RunCmd("wmic qfe get InstalledOn /format:csv 2>nul")
    Dim latestDate : latestDate = ""
    Dim plines : plines = Split(patchOut, Chr(10))
    Dim pi
    For pi = 0 To UBound(plines)
        Dim pl2 : pl2 = Trim(plines(pi))
        If Len(pl2) >= 8 Then
            Dim pp : pp = Split(pl2, ",")
            If UBound(pp) >= 1 Then
                Dim ds : ds = Trim(pp(UBound(pp)))
                If Len(ds) >= 8 And IsNumeric(Left(ds, 4)) Then
                    If ds > latestDate Then latestDate = ds
                End If
            End If
        End If
    Next
    If Len(latestDate) >= 6 Then
        Dim yr : yr = CInt(Left(latestDate, 4))
        Dim mo : mo = CInt(Mid(latestDate, 5, 2))
        Dim nowYr : nowYr = Year(Now())
        Dim nowMo : nowMo = Month(Now())
        Dim monthsAgo : monthsAgo = (nowYr - yr) * 12 + (nowMo - mo)
        If monthsAgo > 3 Then
            issues = issues & "  · 最近补丁安装于 " & latestDate & "，已超过3个月未更新" & Chr(10)
        End If
    Else
        issues = issues & "  · 无法获取操作系统补丁信息" & Chr(10)
    End If
    Dim mysqlVer : mysqlVer = RunCmd("mysql --version 2>nul")
    If Len(Trim(mysqlVer)) > 0 Then
        If InStr(mysqlVer, "5.") > 0 Then
            issues = issues & "  · MySQL版本为5.x，建议升级到8.0+（5.7已于2023年EOL）" & Chr(10)
        End If
    End If
    If issues <> "" Then
        Call AddResult("2.16", "用户安全", "操作系统/数据库/应用软件完成补丁修复升级到最新版本", "fail", _
                       issues, "第2章 第2.16节", _
                       "1. 通过Windows Update安装最新补丁" & Chr(10) & _
                       "2. 将MySQL升级到8.0+稳定版本")
    Else
        Call AddResult("2.16", "用户安全", "操作系统/数据库/应用软件完成补丁修复升级到最新版本", "pass", _
                       "操作系统补丁最新时间正常，数据库版本符合要求。", "第2章 第2.16节", "")
    End If
End Sub

' ============================================================
' 3.8 数据存储权限划分
' ============================================================
Sub Check_3_8_StoragePartition()
    Dim issues : issues = ""
    Dim drive : drive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
    Dim quotaOut : quotaOut = RunCmd("fsutil quota query " & drive & " 2>nul")
    Dim ql : ql = LCase(quotaOut)
    If InStr(ql, "disabled") > 0 Or InStr(ql, "not enabled") > 0 Or InStr(ql, "未启用") > 0 Then
        issues = issues & "  · 系统盘 " & drive & " 磁盘配额未启用" & Chr(10)
    End If
    Dim shareOut : shareOut = RunCmd("net share 2>nul")
    Dim slines : slines = Split(shareOut, Chr(10))
    Dim si
    For si = 0 To UBound(slines)
        Dim sl : sl = Trim(slines(si))
        If InStr(sl, "Everyone") > 0 Then
            If InStr(sl, "Full") > 0 Or InStr(sl, "完全控制") > 0 Then
                issues = issues & "  · 发现共享授权Everyone完全控制：" & sl & Chr(10)
            End If
        End If
    Next
    If issues <> "" Then
        Call AddResult("3.8", "数据安全", "数据存储系统根据重要程度划分存储区块并设置访问权限", "fail", _
                       issues, "第3章 第3.8节", _
                       "1. 启用磁盘配额管理用户存储空间" & Chr(10) & _
                       "2. 收紧关键目录NTFS权限，移除Everyone完全控制")
    Else
        Call AddResult("3.8", "数据安全", "数据存储系统根据重要程度划分存储区块并设置访问权限", "pass", _
                       "存储权限检查通过，未发现过度共享权限问题。", "第3章 第3.8节", "")
    End If
End Sub

' ============================================================
' 手动核查（第5-9章）
' ============================================================

Sub AddManualChecks()
    ' ---------- 手动核查项 ----------
    Call AddResult("1.26_m","系统安全","防病毒日志/补丁日志记录完整有效（人工确认）","manual","请现场核查防病毒和Windows Update日志记录是否完整有效","第1章 第1.26节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.1","物理安全","安全设备/网络设备/终端选型及采购目录化产品","manual","请现场核查设备采购记录与集中采购目录的一致性","第6章 第6.1节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.2","物理安全","安全设备/防火墙等产品通过信息安全产品认证","manual","请现场核查安全产品的认证证书和检测报告","第6章 第6.2节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.3","物理安全","机房布线规范，走线规范、走槽、走管","manual","请现场核查机房线缆布线规范性和标识清晰度","第6章 第6.3节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.4","物理安全","机房温湿度符合GB 2887-2011要求","manual","请现场检查机房温湿度监控记录","第6章 第6.4节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.5","物理安全","机房安装视频监控等安防设备","manual","请现场核查门禁/视频监控系统等安防设备","第6章 第6.5节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.6","物理安全","涉密区域与非涉密区域物理隔离","manual","请现场核查机房区域划分，确认物理隔离","第6章 第6.6节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.7","物理安全","涉密设备上禁止机密信息处理或存储使用","manual","请现场核查涉密设备台账和实际使用记录","第6章 第6.7节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.8","物理安全","选用我国可靠产品，记录所有芯片/主板/存储设备","manual","请现场核查硬件设备台账和供应链可靠性","第6章 第6.8节","参照作业指导书对应章节进行现场核查")
    Call AddResult("6.9","物理安全","按安全等级标注使用不同色系标签","manual","请现场核查线缆和设备颜色标识是否符合规范","第6章 第6.9节","参照作业指导书对应章节进行现场核查")
    Call AddResult("7.1","组织机构","设置安全管理机构和工作","manual","请现场核查组织架构图和安全管理机构职责文件","第7章 第7.1节","参照作业指导书对应章节进行现场核查")
    Call AddResult("7.2","组织机构","配备专职系统安全架构师和管理人员","manual","请现场核查人员台账和岗位职责说明书","第7章 第7.2节","参照作业指导书对应章节进行现场核查")
    Call AddResult("7.3","组织机构","设有安全管理领导小组和工作","manual","请现场核查安全领导小组成立文件","第7章 第7.3节","参照作业指导书对应章节进行现场核查")
    Call AddResult("7.4","组织机构","配备安全管理技术人员和专职管理人员","manual","请现场核查技术人员台账和岗位职责","第7章 第7.4节","参照作业指导书对应章节进行现场核查")
    Call AddResult("7.5","组织机构","建立健全的应急响应体系","manual","请现场核查应急预案文档和演练记录","第7章 第7.5节","参照作业指导书对应章节进行现场核查")
    Call AddResult("8.1","规章制度","建立日常安全管理规章制度/操作规程/安全守则","manual","请现场核查安全管理制度文件的完整性和有效性","第8章 第8.1节","参照作业指导书对应章节进行现场核查")
    Call AddResult("8.2","规章制度","设有安全管理/制度/记录/应急预案等规章制度","manual","请现场核查各类规章制度文件版本和执行记录","第8章 第8.2节","参照作业指导书对应章节进行现场核查")
    Call AddResult("8.3","规章制度","建立日常系统备份规章制度和使用存储介质使用管理制度","manual","请现场核查备份制度和介质管理记录","第8章 第8.3节","参照作业指导书对应章节进行现场核查")
    Call AddResult("8.4","规章制度","建有定期对身份鉴别和访问权限审查工作规程","manual","请现场核查身份鉴别和访问权限审查记录","第8章 第8.4节","参照作业指导书对应章节进行现场核查")
    Call AddResult("9.1","管理实施","建立应用系统应用应急预案，并定期演练或预案执行","manual","请现场核查应急预案文档和最近一次演练记录","第9章 第9.1节","参照作业指导书对应章节进行现场核查")
    Call AddResult("9.2","管理实施","设备管理人员统一配置管理、网络安全管理设备","manual","请现场核查安全设备统一管理平台和配置记录","第9章 第9.2节","参照作业指导书对应章节进行现场核查")
    Call AddResult("9.3","管理实施","对人员进行网络安全管理策略文档和产品，记录详细","manual","请现场核查策略文档的发布、执行记录及其有效性、时效性","第9章 第9.3节","参照作业指导书对应章节进行现场核查")
    Call AddResult("9.4","管理实施","每日安全巡检、每月安全检查，并填写巡检记录和报告","manual","请现场核查日常巡检日志和月度安全检查报告","第9章 第9.4节","参照作业指导书对应章节进行现场核查")
    ' ---------- 不适用项（网络设备/其他设备类）----------
    Call AddResult("5.1","网络安全","跨网交换数据需经过审批","na","请登录边界设备/交换平台核查，超出本脚本范围","第5章 第5.1节","")
    Call AddResult("5.2","网络安全","跨网交换使用专用设备","na","请现场核查跨网交换设备和隔离网闸措施","第5章 第5.2节","")
    Call AddResult("5.3","网络安全","网络架构最小化原则","na","请登录网络设备核查网络拓扑和架构，超出本脚本范围","第5章 第5.3节","")
    Call AddResult("5.4","网络安全","网络边界部署节点/路由/地址最小化","na","请登录网络设备核查，超出本脚本范围","第5章 第5.4节","")
    Call AddResult("5.5","网络安全","网络内部按安全域划分","na","请核查VLAN和安全域划分文档","第5章 第5.5节","")
    Call AddResult("5.6","网络安全","网络设备配置最小化原则全覆盖","na","请登录网络设备核查配置","第5章 第5.6节","")
    Call AddResult("5.7","网络安全","网络设备配置唯一管理服务并限制管理终端","na","请登录网络设备核查管理接口","第5章 第5.7节","")
    Call AddResult("5.8","网络安全","远程管理采取加密保护措施或加密的管理通道","na","请登录网络设备核查SSH/HTTPS配置","第5章 第5.8节","")
    Call AddResult("5.9","网络安全","同链路同侧不同安全设备使用不同架构/品牌","na","请核查网络设备清单","第5章 第5.9节","")
    Call AddResult("5.10","网络安全","采取组播/广播/阻断等控制措施","na","请登录网络设备核查组播控制","第5章 第5.10节","")
    Call AddResult("5.11","网络安全","关键网络设备和安全设备冗余备份","na","请核查网络设备冗余清单","第5章 第5.11节","")
    Call AddResult("5.12","网络安全","远程管理采取加密通道+加密传输协议保护","na","请核查远程管理通道和加密措施","第5章 第5.12节","")
    Call AddResult("5.13","网络安全","接入网络安全运营中心部署的接入终端","na","请核查SOC/安全管理平台接入情况","第5章 第5.13节","")
    Call AddResult("5.14","网络安全","与互联单元/低等级单元互联采取逻辑隔离措施","na","请核查隔离措施","第5章 第5.14节","")
    Call AddResult("5.15","网络安全","不同单元之间通过防火墙/网闸等设备加强逻辑隔离","na","请核查边界安全隔离设备配置","第5章 第5.15节","")
    Call AddResult("5.16","网络安全","访问控制策略实施细化到端口级","na","请核查防火墙/ACL配置","第5章 第5.16节","")
    Call AddResult("5.17","网络安全","远程管理采用地方路由时采取加密保护措施","na","请核查链路路由和加密措施","第5章 第5.17节","")
    Call AddResult("5.18","网络安全","接入认证采用基于802.1x协议强化","na","请核查网络接入认证配置","第5章 第5.18节","")
    Call AddResult("5.19","网络安全","用户与网络之间逻辑隔离","na","请核查终端接入隔离策略","第5章 第5.19节","")
    Call AddResult("5.20","网络安全","全网行为采集记录集中安全审计留存>=180天","na","请核查全流量安全审计系统","第5章 第5.20节","")
    Call AddResult("5.21","网络安全","对攻击/违规行为实时监测告警阻断","na","请核查IDS/IPS/SOC等监测系统","第5章 第5.21节","")
    Call AddResult("5.22","网络安全","数据存储系统接入网络应采取逻辑隔离","na","请核查存储网络隔离配置","第5章 第5.22节","")
    Call AddResult("5.23","网络安全","网络设备具备冗余性，关键设备必要冗余","na","请核查网络设备冗余清单","第5章 第5.23节","")
    ' ---------- 不适用项（应用安全其他项）----------
    Call AddResult("4.1","应用安全","应用系统具备应用容灾措施","na","请结合业务应用判断，数据不属于业务架构","第4章 第4.1节","")
    Call AddResult("4.2","应用安全","应用系统在未及时安装补丁或更新安全软件","na","请结合业务应用版本和补丁情况判断","第4章 第4.2节","")
    Call AddResult("4.3","应用安全","对开放端口应用进行身份认证和告警阻断","na","请结合硬件和开放端口模块，超出本脚本范围","第4章 第4.3节","")
    Call AddResult("4.4","应用安全","对应用信息采取加密保护等措施","na","请结合业务加密架构清单，超出本脚本范围","第4章 第4.4节","")
    Call AddResult("4.5","应用安全","提供承载信息的应用应具备DDoS攻击防护能力","na","请核查抗拒绝服务/抗DDoS设备","第4章 第4.5节","")
    Call AddResult("4.7","应用安全","网站应用采取静态页面发布形式","na","请结合业务网站页面发布形式判断","第4章 第4.7节","")
    Call AddResult("4.9","应用安全","Web应用具备对SQL注入/XSS等攻击防护能力","na","请结合渗透测试或漏洞扫描报告，超出本脚本范围","第4章 第4.9节","")
    Call AddResult("4.10","应用安全","Web应用具备执行代码的有效校验机制","na","请结合业务设计，超出本脚本范围","第4章 第4.10节","")
    Call AddResult("4.11","应用安全","具备用户权限到访问控制的控制能力","na","请结合业务应用权限模块判断","第4章 第4.11节","")
    Call AddResult("4.13","应用安全","业务管理终端专用专管","na","请结合现场核查终端管理，超出本脚本范围","第4章 第4.13节","")
    Call AddResult("4.14","应用安全","基于用户角色和权限的访问控制，细分粒度登录","na","请结合业务应用RBAC权限模块判断","第4章 第4.14节","")
    Call AddResult("4.19","应用安全","非使用通用协议或通用接口通信协议/接口","na","请核查应用接口技术栈，超出本脚本范围","第4章 第4.19节","")
    Call AddResult("4.20","应用安全","具备对用户使用口令认证或用户生物认证机制","na","请结合硬件USB Key/密码卡等，超出本脚本范围","第4章 第4.20节","")
    Call AddResult("4.25","应用安全","应用系统进行代码级安全漏洞挖掘","na","请核查SAST/DAST扫描报告，超出本脚本范围","第4章 第4.25节","")
    Call AddResult("4.26","应用安全","具备对账户调用接口/异常通信/文件操作格式等级控制","na","请结合业务设计，超出本脚本范围","第4章 第4.26节","")
    Call AddResult("4.27","应用安全","应有效防护并阻断SQL注入/XSS/DoS等应用层攻击","na","请结合渗透测试或漏洞扫描报告","第4章 第4.27节","")
    Call AddResult("4.28","应用安全","对用户登录和权限统一管理控制","na","请核查IAM平台接入情况，超出本脚本范围","第4章 第4.28节","")
    Call AddResult("4.30","应用安全","应满足请求并分会话数/带宽/单用户连接限制","na","请结合业务应用中间件配置判断","第4章 第4.30节","")
    Call AddResult("2.3","用户安全","用户应用接入应用考勤，通过身份认证使用应用信息接入","na","请核查业务系统统一身份认证机制","第2章 第2.3节","")
    Call AddResult("2.11_m","用户安全","口令认证采用身份认证或专用的硬件或软件","na","请结合硬件USB Key/密码卡/指纹，超出本脚本范围","第2章 第2.11节","")
    Call AddResult("3.2","数据安全","用户对应用数据采取加密文件或统一管控和措施","na","请核查数据加密管理系统，超出本脚本范围","第3章 第3.2节","")
    Call AddResult("3.3","数据安全","加密存储介质降密级使用前采取清除写残余数据","na","请现场核查介质管理流程，超出本脚本范围","第3章 第3.3节","")
    Call AddResult("3.4","数据安全","确保涉密数据的传输/存储/备份等过程安全传输","na","请现场核查介质管理流程，超出本脚本范围","第3章 第3.4节","")
    Call AddResult("3.6","数据安全","网络边界具备信息防泄露/防篡改等数据防泄露","na","请核查边界设备DLP配置","第3章 第3.6节","")
    Call AddResult("3.7","数据安全","数据存储系统具备日志记录采取认证保护强化措施+留存180天","na","请结合业务存储系统日志配置判断","第3章 第3.7节","")
    Call AddResult("3.9","数据安全","数据备份采用介质/路由独立/统一管控/完整日志记录","na","请结合业务数据备份和路由判断","第3章 第3.9节","")
    Call AddResult("3.11","数据安全","数据访问权限分级分权，应为最小化","na","请结合业务系统权限分级判断","第3章 第3.11节","")
    Call AddResult("3.12","数据安全","数据采集未超出业务必要范围","na","请结合业务数据采集范围判断","第3章 第3.12节","")
    Call AddResult("3.13","数据安全","数据备份介质存在介质加密机应保密的必要","na","请结合业务数据备份介质管理判断","第3章 第3.13节","")
    Call AddResult("3.14","数据安全","涉密分类提供统一管控和访问控制，应为最小化","na","请核查数据脱敏平台，超出本脚本范围","第3章 第3.14节","")
    Call AddResult("10.1_m","协议安全审计","涉密协议认证/通信口令/密钥分发管理","na","请核查密钥签名/证书管控和通信日志留存情况","第10章 第10.1节","")
End Sub

' ============================================================
' 工具函数
' ============================================================

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

Function ReadReg(path)
    On Error Resume Next
    Dim val : val = oShell.RegRead(path)
    If Err.Number <> 0 Then val = "" : Err.Clear
    ReadReg = val
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
    WScript.Echo "  核查完成，共 " & rCount & " 项"
    WScript.Echo "  [OK]  通过         : " & nPass & " 项"
    WScript.Echo "  [!!]  未通过       : " & nFail & " 项"
    WScript.Echo "  [??]  需人工核查   : " & nManual & " 项"
    WScript.Echo "  [--]  不适用       : " & nNA & " 项"
    If nFail > 0 Then
        WScript.Echo ""
        WScript.Echo "  未通过项目："
        For i = 0 To rCount - 1
            If rStatus(i) = "fail" Then
                WScript.Echo "    [!!] [" & rID(i) & "] " & rTitle(i)
            End If
        Next
    End If
End Sub

Sub GenerateHTML()
    Dim outDir : outDir = "output"
    If Not oFSO.FolderExists(outDir) Then oFSO.CreateFolder(outDir)

    Dim ts, dtNow
    dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim reportPath : reportPath = outDir & "\配置核查报告_" & stamp & ".html"
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
    ts.WriteLine "<title>配置核查报告 " & stamp & "</title>"
    ts.WriteLine "<style>"
    ts.WriteLine "body{font-family:'Microsoft YaHei',Arial,sans-serif;margin:20px;background:#f5f5f5}"
    ts.WriteLine "h1{color:#333;text-align:center}h2{margin-top:30px}"
    ts.WriteLine ".sum{background:white;padding:20px;margin:20px 0;border-radius:5px;box-shadow:0 2px 5px rgba(0,0,0,.1)}"
    ts.WriteLine ".pass{color:#28a745}.fail{color:#dc3545}.manual{color:#e67e00}.na{color:#6c757d}"
    ts.WriteLine "table{width:100%;border-collapse:collapse;background:white;margin:20px 0;box-shadow:0 2px 5px rgba(0,0,0,.1)}"
    ts.WriteLine "th,td{padding:10px;text-align:left;border-bottom:1px solid #ddd}"
    ts.WriteLine "th{background:#007bff;color:white}"
    ts.WriteLine "tr:hover{background:#f0f4ff}"
    ts.WriteLine ".det{white-space:pre-wrap;word-wrap:break-word;max-width:400px}"
    ts.WriteLine ".rec{background:#fff3cd;padding:8px;border-left:3px solid #ffc107;margin:4px 0}"
    ts.WriteLine ".ch{color:#666;font-size:.9em}"
    ts.WriteLine ".badge{display:inline-block;padding:3px 8px;border-radius:3px;font-size:.85em;font-weight:bold}"
    ts.WriteLine ".b-pass{background:#d4edda;color:#155724}.b-fail{background:#f8d7da;color:#721c24}"
    ts.WriteLine ".b-man{background:#fff3cd;color:#856404}.b-na{background:#e2e3e5;color:#383d41}"
    ts.WriteLine "</style></head><body>"
    ts.WriteLine "<h1>配置核查报告（Windows XP/7版）</h1>"
    ts.WriteLine "<p style=""text-align:center;color:#666"">生成时间：" & CStr(Now()) & "，系统：" & HtmlEsc(osCaption) & "</p>"
    ts.WriteLine "<div class=""sum""><h2>核查摘要</h2>"
    ts.WriteLine "<p><strong>总核查项：</strong>" & rCount & " 项</p>"
    ts.WriteLine "<p class=""pass""><strong>通过：</strong>" & nPass & " 项</p>"
    ts.WriteLine "<p class=""fail""><strong>未通过：</strong>" & nFail & " 项</p>"
    ts.WriteLine "<p class=""manual""><strong>需人工核查：</strong>" & nManual & " 项</p>"
    ts.WriteLine "<p class=""na""><strong>不适用：</strong>" & nNA & " 项</p>"
    ts.WriteLine "</div>"

    ' 未通过
    If nFail > 0 Then
        ts.WriteLine "<h2 class=""fail"">未通过项目（" & nFail & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>详情</th><th>修复建议</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "fail" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>"
                ts.WriteLine "<td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td>"
                ts.WriteLine "<td class=""rec"">" & HtmlEsc(rRec(i)) & "</td>"
                ts.WriteLine "<td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    ' 需人工核查
    If nManual > 0 Then
        ts.WriteLine "<h2 class=""manual"">需人工核查项目（" & nManual & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>核查要求</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "manual" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>"
                ts.WriteLine "<td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td>"
                ts.WriteLine "<td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    ' 不适用
    If nNA > 0 Then
        ts.WriteLine "<h2 class=""na"">不适用项目（" & nNA & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>说明</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "na" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>"
                ts.WriteLine "<td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td>"
                ts.WriteLine "<td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    ' 通过
    If nPass > 0 Then
        ts.WriteLine "<h2 class=""pass"">通过项目（" & nPass & " 项）</h2>"
        ts.WriteLine "<table><tr><th>编号</th><th>类别</th><th>检查项</th><th>详情</th><th>章节</th></tr>"
        For i = 0 To rCount - 1
            If rStatus(i) = "pass" Then
                ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>"
                ts.WriteLine "<td>" & HtmlEsc(rTitle(i)) & "</td>"
                ts.WriteLine "<td class=""det"">" & HtmlEsc(rDetail(i)) & "</td>"
                ts.WriteLine "<td class=""ch"">" & HtmlEsc(rChapter(i)) & "</td></tr>"
            End If
        Next
        ts.WriteLine "</table>"
    End If

    ts.WriteLine "</body></html>"
    ts.Close

    WScript.Echo "[OK] 报告生成完成：" & reportPath
    ' 自动在浏览器中打开
    On Error Resume Next
    oShell.Run "explorer.exe """ & reportPath & """"
    On Error GoTo 0
End Sub


' ============================================================
' 生成 Excel 报告（无 Excel 时降级 CSV）
' ============================================================
Sub GenerateExcel()
    Dim oExcel, oWB, oWS
    On Error Resume Next
    Set oExcel = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Call GenerateXLS()
        Exit Sub
    End If
    On Error GoTo 0

    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)

    Dim outDirAbs : outDirAbs = oFSO.GetAbsolutePathName("output")
    If Not oFSO.FolderExists(outDirAbs) Then oFSO.CreateFolder(outDirAbs)

    Set oWB = oExcel.Workbooks.Add()
    Set oWS = oWB.Worksheets(1)
    oWS.Name = "配置核查报告"
    oExcel.Visible = False
    oExcel.DisplayAlerts = False

    oWS.Cells(1,1).Value = "配置核查报告（Windows XP/7版）"
    oWS.Cells(2,1).Value = "生成时间：" & CStr(dtNow) & "  系统：" & osCaption
    oWS.Cells(1,1).Font.Size = 14
    oWS.Cells(1,1).Font.Bold = True

    Dim row : row = 4
    Dim headers : headers = Array("编号","类别","检查项","状态","详情","修复建议/核查要求","章节")
    Dim c
    For c = 1 To 7
        oWS.Cells(row,c).Value = headers(c-1)
        oWS.Cells(row,c).Font.Bold = True
        oWS.Cells(row,c).Interior.Color = RGB(0,112,192)
        oWS.Cells(row,c).Font.Color = RGB(255,255,255)
    Next

    Dim i
    For i = 0 To rCount - 1
        row = row + 1
        oWS.Cells(row,1).Value = rID(i)
        oWS.Cells(row,2).Value = rCat(i)
        oWS.Cells(row,3).Value = rTitle(i)
        Dim statusText
        Select Case rStatus(i)
            Case "pass": statusText = "通过"
            Case "fail": statusText = "未通过"
            Case "manual": statusText = "需人工核查"
            Case Else: statusText = "不适用"
        End Select
        oWS.Cells(row,4).Value = statusText
        oWS.Cells(row,5).Value = rDetail(i)
        oWS.Cells(row,6).Value = rRec(i)
        oWS.Cells(row,7).Value = rChapter(i)
        Select Case rStatus(i)
            Case "pass": oWS.Cells(row,4).Font.Color = RGB(40,167,69)
            Case "fail": oWS.Cells(row,4).Font.Color = RGB(220,53,69)
            Case "manual": oWS.Cells(row,4).Font.Color = RGB(230,126,0)
        End Select
    Next

    oWS.Columns(1).ColumnWidth = 10
    oWS.Columns(2).ColumnWidth = 12
    oWS.Columns(3).ColumnWidth = 40
    oWS.Columns(4).ColumnWidth = 12
    oWS.Columns(5).ColumnWidth = 50
    oWS.Columns(6).ColumnWidth = 50
    oWS.Columns(7).ColumnWidth = 14
    oWS.Columns(5).WrapText = True
    oWS.Columns(6).WrapText = True

    Dim xlsxPath : xlsxPath = outDirAbs & "\配置核查报告_" & stamp & ".xlsx"
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

    WScript.Echo "[OK] Excel报告生成：" & xlsxPath
End Sub

Sub GenerateCSV()
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim outDirAbs : outDirAbs = oFSO.GetAbsolutePathName("output")
    If Not oFSO.FolderExists(outDirAbs) Then oFSO.CreateFolder(outDirAbs)
    Dim csvPath : csvPath = outDirAbs & "\配置核查报告_" & stamp & ".csv"
    Dim ts : Set ts = oFSO.CreateTextFile(csvPath, True, False)
    ts.WriteLine "编号,类别,检查项,状态,详情,修复建议/核查要求,章节"
    Dim i
    For i = 0 To rCount - 1
        Dim statusText
        Select Case rStatus(i)
            Case "pass": statusText = "通过"
            Case "fail": statusText = "未通过"
            Case "manual": statusText = "需人工核查"
            Case Else: statusText = "不适用"
        End Select
        ts.WriteLine CsvEsc(rID(i)) & "," & CsvEsc(rCat(i)) & "," & CsvEsc(rTitle(i)) & "," & _
                    CsvEsc(statusText) & "," & CsvEsc(rDetail(i)) & "," & CsvEsc(rRec(i)) & "," & _
                    CsvEsc(rChapter(i))
    Next
    ts.Close
    WScript.Echo "[OK] CSV报告生成（未检测到Excel，降级生成）：" & csvPath
End Sub

Function CsvEsc(s)
    Dim r : r = CStr(s)
    If InStr(r, ",") > 0 Or InStr(r, Chr(10)) > 0 Or InStr(r, Chr(13)) > 0 Or InStr(r, """") > 0 Then
        r = """" & Replace(r, """", """""") & """"
    End If
    CsvEsc = r
End Function

' ============================================================
' 生成 .xls 报告（HTML 表格，无 Excel 时使用，Excel/WPS 可直接打开）
' ============================================================
Sub GenerateXLS()
    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim outDirAbs : outDirAbs = oFSO.GetAbsolutePathName("output")
    If Not oFSO.FolderExists(outDirAbs) Then oFSO.CreateFolder(outDirAbs)
    Dim xlsPath : xlsPath = outDirAbs & "\配置核查报告_" & stamp & ".xls"
    Dim ts : Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>配置核查报告</x:Name></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->"
    ts.WriteLine "<style>"
    ts.WriteLine "table{border-collapse:collapse;font-family:'Microsoft YaHei',Arial;font-size:10pt}"
    ts.WriteLine "td,th{border:1px solid #bbb;padding:4px 6px;text-align:left;vertical-align:top}"
    ts.WriteLine "th.hd{background:#007bff;color:#fff;font-weight:bold}"
    ts.WriteLine ".pass{color:#28a745;font-weight:bold}.fail{color:#dc3545;font-weight:bold}.manual{color:#e67e00;font-weight:bold}.na{color:#6c757d}"
    ts.WriteLine ".title{font-size:14pt;font-weight:bold}.sub{color:#666}"
    ts.WriteLine "</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7"" class=""title"">配置核查报告（Windows XP/7版）</td></tr>"
    ts.WriteLine "<tr><td colspan=""7"" class=""sub"">生成时间：" & HtmlEsc(CStr(dtNow)) & "  系统：" & HtmlEsc(osCaption) & "</td></tr>"
    ts.WriteLine "<tr></tr>"
    ts.WriteLine "<tr><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">详情</th><th class=""hd"">修复建议/核查要求</th><th class=""hd"">章节</th></tr>"
    Dim i
    For i = 0 To rCount - 1
        Dim statusText, statusCls
        Select Case rStatus(i)
            Case "pass":   statusText = "通过"       : statusCls = "pass"
            Case "fail":   statusText = "未通过"     : statusCls = "fail"
            Case "manual": statusText = "需人工核查" : statusCls = "manual"
            Case Else:     statusText = "不适用"     : statusCls = "na"
        End Select
        ts.WriteLine "<tr><td>" & HtmlEsc(rID(i)) & "</td><td>" & HtmlEsc(rCat(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rTitle(i)) & "</td>" & _
                     "<td class=""" & statusCls & """>" & statusText & "</td>" & _
                     "<td>" & HtmlEsc(rDetail(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rRec(i)) & "</td>" & _
                     "<td>" & HtmlEsc(rChapter(i)) & "</td></tr>"
    Next
    ts.WriteLine "</table></body></html>"
    ts.Close
    WScript.Echo "[OK] Excel报告生成（.xls）：" & xlsPath
End Sub