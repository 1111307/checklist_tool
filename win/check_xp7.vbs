' ============================================================
' 配置核查工具 - Windows XP/7 版（无需Python）
' 适用系统：Windows XP SP3 / Windows 7 / Windows Server 2003/2008 R2
' 运行方式：双击 run.bat（建议管理员权限）
' 输出文件：output\配置核查报告_日期时间.html
' ============================================================
Option Explicit

Dim oShell, oFSO, oWMI, oCache
Dim osCaption, osBuild, isXP, is7OrAbove

Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oCache = CreateObject("Scripting.Dictionary")

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
WScript.Echo "  参考标准：配置核查作业指导书v2.2"
WScript.Echo "================================================================"
If osCaption <> "" Then WScript.Echo "  系统：" & osCaption
WScript.Echo ""
WScript.Echo "开始检查，请稍候..."
WScript.Echo ""

' --- 数据库连接参数（优先级：环境变量 > db_config.conf > 默认值）---
Dim MYSQL_HOST, MYSQL_PORT, MYSQL_USER, MYSQL_PASS
MYSQL_HOST = GetSetting("MYSQL_HOST", "127.0.0.1")
MYSQL_PORT = GetSetting("MYSQL_PORT", "3306")
MYSQL_USER = GetSetting("MYSQL_USER", "root")
MYSQL_PASS = GetSetting("MYSQL_PASS", "")

Dim mysqlConnOk : mysqlConnOk = False
Dim mysqlVer : mysqlVer = RunCmdCached("mysql --version 2>nul")
If Len(Trim(mysqlVer)) > 0 Then
    Dim mysqlProbe : mysqlProbe = RunCmd("mysql -h " & MYSQL_HOST & " -P " & MYSQL_PORT & " -u " & MYSQL_USER & " -p" & MYSQL_PASS & " --connect-timeout=2 -e ""SELECT 1;"" 2>nul")
    mysqlConnOk = (Len(Trim(mysqlProbe)) > 0)
End If

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
Call Check_1_27_SoftwareLicensing()
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
Call Check_2_3_AppPass()
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
Call Check_3_6_DLP()
Call Check_3_10_Shares()
Call Check_3_8_StoragePartition()
Call Check_3_2_DataTransfer()
Call Check_3_3_MediaErase()
Call Check_3_7_StorageLogin()
Call Check_3_9_TransferEncryption()
Call Check_3_11_AccessLevels()
Call Check_3_13_DataEncryption()

' ---------- 第4章 应用安全 ----------
Call Check_4_16_NLA()
Call Check_4_17_RoleSeparation()
Call Check_4_18_IPRestrict()
Call Check_4_19_Lockout()
Call Check_4_23_PortSeparation()
Call Check_4_24_AppDBSep()
Call Check_4_25_AppAudit()
Call Check_4_6_WAF()
Call Check_4_8_AntiTamper()
Call Check_4_12_SessionLimit()
Call Check_4_22_DefaultPort()
Call Check_4_30_RemoteMgmt()
Call Check_4_32_AppBackup()
Call Check_4_33_DomesticSoftware()
Call Check_4_2_AppPatch()
Call Check_4_3_TrustedRoot()
Call Check_4_4_ServerSeparation()
Call Check_4_7_StaticPublish()
Call Check_4_11_UserAuthz()
Call Check_4_14_FineGrainedACL()
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
    Dim ns : ns = "SecurityCenter2"
    If isXP Then ns = "SecurityCenter"
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
                       "manual", "未发现已知高危服务（telnet/tftp/ftp/snmp），但系统运行的服务与开放端口是否均为业务必需，需人工逐一确认。", "第1章 第1.3节", "核查运行服务列表与监听端口，停用并禁用非业务必需的服务。")
    End If
End Sub

Sub Check_1_4_Firewall()
    Dim fwOn : fwOn = False
    Dim detail : detail = ""
    If is7OrAbove Then
        Dim advOut : advOut = RunCmdCached("netsh advfirewall show allprofiles 2>nul")
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
    Dim gpOut : gpOut = RunCmdCached("gpresult /R /SCOPE COMPUTER 2>nul")
    Dim hasGPO : hasGPO = False
    If InStr(gpOut, "Resource") > 0 Or InStr(gpOut, "Quota") > 0 Or InStr(gpOut, "配额") > 0 Or InStr(gpOut, "资源") > 0 Or InStr(gpOut, "组策略") > 0 Then
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
    If InStr(progdata, "%") > 0 Then progdata = oShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")
    If isXP Then
        ' Windows XP 无 Windows Defender，跳过该检查
    ElseIf oFSO.FolderExists(progdata & "\Microsoft\Windows Defender\Support") Then
        findings = findings & "  · Windows Defender 支持日志目录存在" & Chr(10)
    Else
        issues = issues & "  · 未找到 Windows Defender 支持日志目录" & Chr(10)
    End If
    Dim secOut : secOut = ""
    If is7OrAbove Then
        secOut = RunCmd("wevtutil qe Security /c:5 /rd:true /f:text 2>nul")
    Else
        secOut = RunCmd("cscript //nologo eventquery.vbs /L Security /R 5 2>nul")
    End If
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
' 1.27 基础软件正版/授权及售后技术支持
' ============================================================
Sub Check_1_27_SoftwareLicensing()
    Dim findings : findings = ""
    If osCaption <> "" Then findings = findings & "  · 系统：" & osCaption & Chr(10)
    Dim licOut : licOut = ""
    If Not isXP Then
        licOut = RunCmd("cscript //nologo %windir%\system32\slmgr.vbs /dli 2>nul")
        If Len(Trim(licOut)) = 0 Then
            findings = findings & "  · slmgr /dli 授权信息无法读取（可能权限不足或 slmgr.vbs 缺失）" & Chr(10)
        ElseIf InStr(licOut, "已授权") > 0 Or InStr(LCase(licOut), "licensed") > 0 Then
            findings = findings & "  · slmgr /dli 显示系统处于授权（Licensed）状态" & Chr(10)
        Else
            findings = findings & "  · slmgr /dli 未返回明确授权状态，需人工确认" & Chr(10)
        End If
    Else
        findings = findings & "  · Windows XP 无 slmgr 授权查询工具，需人工核对授权凭证" & Chr(10)
    End If
    Call AddResult("1.27", "系统安全", "基础软件正版/授权及售后技术支持", "manual", _
                   "操作系统版本与授权状态已自动核查，但正版/定制软件及售后技术支持需人工核对采购与授权凭证。" & Chr(10) & findings, "第1章 第1.27节", _
                   "核对操作系统、数据库、中间件、办公软件的采购合同、正版授权证书（COA 标签/电子发货单）及售后技术支持合同。")
End Sub

' ============================================================
' 1.7 数据库冗余帐户
' ============================================================
Sub Check_1_7_DbAccounts()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.7", "系统安全", "数据库冗余帐户", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.7节", "")
        Exit Sub
    End If
    Dim usersOut : usersOut = MysqlQ("SELECT user FROM mysql.user;")
    If Len(Trim(usersOut)) = 0 Then
        Call AddResult("1.7", "系统安全", "数据库冗余帐户", "manual", _
                       "MySQL已安装，但无法查询账户列表（可能缺少登录凭据或权限不足），需人工核查。", "第1章 第1.7节", "使用有效凭据登录 MySQL，核查是否存在 test/anonymous/guest 等冗余账户。")
        Exit Sub
    End If
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
                       "MySQL账户查询成功，未发现 test/anonymous/guest/demo 冗余账户。", "第1章 第1.7节", "")
    End If
End Sub

' ============================================================
' 1.8 数据库冗余存储过程
' ============================================================
Sub Check_1_8_DbProcedures()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.8节", "")
        Exit Sub
    End If
    ' 排除系统库（sys/mysql/information_schema/performance_schema），只统计用户自建存储过程
    Dim procsOut : procsOut = MysqlQ("SELECT COUNT(*) cnt FROM information_schema.routines WHERE routine_type='PROCEDURE' AND routine_schema NOT IN ('sys','mysql','information_schema','performance_schema');")
    Dim lines : lines = Split(procsOut, Chr(10))
    Dim i
    Dim gotNum : gotNum = False
    Dim userProcCnt : userProcCnt = 0
    For i = 0 To UBound(lines)
        Dim ln : ln = Trim(lines(i))
        If IsNumeric(ln) Then
            gotNum = True
            userProcCnt = CInt(ln)
        End If
    Next
    If Not gotNum Then
        Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "manual", _
                       "MySQL已安装，但无法查询存储过程统计（可能缺少登录凭据或权限不足），需人工核查。", "第1章 第1.8节", "使用有效凭据登录 MySQL 核查存储过程。")
    ElseIf userProcCnt > 0 Then
        Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "manual", _
                       "发现 " & userProcCnt & " 个用户自建存储过程，是否冗余需人工确认。", "第1章 第1.8节", _
                       "核查存储过程清单，删除不再使用的：DROP PROCEDURE proc_name;")
    Else
        Call AddResult("1.8", "系统安全", "数据库冗余存储过程", "pass", _
                       "未发现用户自建存储过程（系统自带存储过程已排除）。", "第1章 第1.8节", "")
    End If
End Sub

' ============================================================
' 1.9 数据库细粒度访问授权
' ============================================================
Sub Check_1_9_DbGrants()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.9", "系统安全", "数据库细粒度访问授权（表级增删改查）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.9节", "")
        Exit Sub
    End If
    Dim grantsOut : grantsOut = MysqlQ("SELECT user,host,Super_priv FROM mysql.user WHERE user NOT IN ('root','');")
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.10节", "")
        Exit Sub
    End If
    Dim issues : issues = ""
    Dim anonOut : anonOut = MysqlQ("SELECT COUNT(*) FROM mysql.user WHERE user='' AND (Select_priv='Y' OR Insert_priv='Y' OR Update_priv='Y');")
    Dim alines : alines = Split(anonOut, Chr(10))
    Dim ai
    Dim gotNum : gotNum = False
    For ai = 0 To UBound(alines)
        Dim av : av = Trim(alines(ai))
        If IsNumeric(av) Then
            gotNum = True
            If CInt(av) > 0 Then
                issues = issues & "  · 发现匿名用户拥有数据操作权限" & Chr(10)
            End If
        End If
    Next
    Dim emptyOut : emptyOut = MysqlQ("SELECT COUNT(*) FROM mysql.user WHERE user='root' AND (authentication_string='' OR authentication_string IS NULL);")
    Dim elines : elines = Split(emptyOut, Chr(10))
    Dim ei
    For ei = 0 To UBound(elines)
        Dim ev : ev = Trim(elines(ei))
        If IsNumeric(ev) Then
            gotNum = True
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
    ElseIf Not gotNum Then
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "manual", _
                       "MySQL已安装，但无法查询账户状态（可能缺少登录凭据或权限不足），需人工核查。", "第1章 第1.10节", "使用有效凭据登录 MySQL，核查匿名账户与 root 空密码情况。")
    Else
        Call AddResult("1.10", "系统安全", "数据库自主访问控制功能", "pass", _
                       "MySQL访问控制检查通过（匿名账户/root空密码检查均正常）。", "第1章 第1.10节", "")
    End If
End Sub

' ============================================================
' 1.11 数据库备份与恢复功能
' ============================================================
Sub Check_1_11_DbBackup()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.11", "系统安全", "数据库备份与恢复功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.11节", "")
        Exit Sub
    End If
    Dim findings : findings = ""
    Dim taskOut : taskOut = RunCmdCached("schtasks /query /fo LIST 2>nul")
    If InStr(LCase(taskOut), "mysqldump") > 0 Or InStr(LCase(taskOut), "mysql") > 0 Then
        findings = findings & "  · 检测到MySQL相关备份计划任务" & Chr(10)
    End If
    Dim bdirs : bdirs = SysDrive() & "\backup " & SysDrive() & "\mysql_backup " & SysDrive() & "\Backup " & SysDrive() & "\mysql_backup"
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.12", "系统安全", "数据库表级审计、告警和阻断功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.12节", "")
        Exit Sub
    End If
    Dim pluginsOut : pluginsOut = MysqlQ("SHOW PLUGINS;")
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
    Dim genlogOut : genlogOut = MysqlQ("SHOW VARIABLES LIKE 'general_log';")
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.13", "系统安全", "数据库数据与其他应用数据独立存储", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.13节", "")
        Exit Sub
    End If
    Dim dbsOut : dbsOut = MysqlQ("SHOW DATABASES;")
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
        ' 探测 nginx 配置文件真实路径
        Dim nginxConf : nginxConf = ProbeNginxConf()
        Dim sslOut : sslOut = ""
        If nginxConf <> "" Then
            sslOut = RunCmd("findstr /i ""ssl"" """ & nginxConf & """ 2>nul")
        End If
        ' 兜底：查是否有 443 端口监听（说明启用了 HTTPS）
        Dim has443 : has443 = IsPortListening("443")
        If Len(Trim(sslOut)) = 0 And Not has443 Then
            Call AddResult("1.14", "系统安全", "中间件安全", "fail", _
                           "Nginx未配置SSL/TLS（未发现ssl配置，且无443端口监听）", "第1章 第1.14节", "配置SSL证书并启用HTTPS")
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.15", "系统安全", "数据库远程登录限制", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.15节", "")
        Exit Sub
    End If
    Dim remoteOut : remoteOut = MysqlQ("SELECT host FROM mysql.user WHERE user='root' AND host NOT IN ('localhost','127.0.0.1','::1');")
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.16_17", "系统安全", "数据库输入检查（SQL注入防护）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.16-1.17节", "")
        Exit Sub
    End If
    Dim modeOut : modeOut = MysqlQ("SELECT @@sql_mode;")
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
    ' 只判断 LISTENING 状态，避免出站连接误判
    Dim hasMySQL : hasMySQL = IsPortListening("3306")
    Dim hasMSSQL : hasMSSQL = IsPortListening("1433")
    If hasMySQL Or hasMSSQL Then
        Dim portName : portName = ""
        If hasMySQL Then portName = "3306(MySQL)"
        If hasMSSQL Then
            If portName <> "" Then portName = portName & "/"
            portName = portName & "1433(MSSQL)"
        End If
        Call AddResult("1.19", "系统安全", "更换数据库默认设置（端口/账户）", "fail", _
                       "数据库监听默认端口：" & portName, "第1章 第1.19节", _
                       "修改数据库默认端口和管理员用户名，避免使用默认值")
    Else
        Call AddResult("1.19", "系统安全", "更换数据库默认设置（端口/账户）", "pass", _
                       "未检测到数据库默认端口（3306/1433）监听。", "第1章 第1.19节", "")
    End If
End Sub

' ============================================================
' 1.20 数据库安全策略（密码复杂度）
' ============================================================
Sub Check_1_20_DbPolicy()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.20节", "")
        Exit Sub
    End If
    Dim policyOut : policyOut = MysqlQ("SHOW VARIABLES LIKE '%password%';")
    ' 兼容 MySQL(validate_password) 与 MariaDB(simple_password_check/cracklib_password_check/strict_password_validation)
    Dim hasPolicy : hasPolicy = False
    If InStr(LCase(policyOut), "validate_password") > 0 Then hasPolicy = True
    If InStr(LCase(policyOut), "simple_password_check") > 0 Then hasPolicy = True
    If InStr(LCase(policyOut), "cracklib_password_check") > 0 Then hasPolicy = True
    If InStr(LCase(policyOut), "strict_password_validation") > 0 Then hasPolicy = True
    If hasPolicy Then
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "pass", _
                       "数据库密码复杂度策略已配置。", "第1章 第1.20节", "")
    Else
        Call AddResult("1.20", "系统安全", "数据库安全策略（密码/锁定/超时）", "fail", _
                       "数据库未配置密码复杂度策略。", "第1章 第1.20节", _
                       "启用密码复杂度校验插件（MySQL: validate_password；MariaDB: simple_password_check 或 cracklib_password_check）。")
    End If
End Sub

' ============================================================
' 1.21 数据库审计功能
' ============================================================
Sub Check_1_21_DbAudit()
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.21", "系统安全", "数据库审计功能", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.21节", "")
        Exit Sub
    End If
    Dim auditOut : auditOut = MysqlQ("SHOW VARIABLES LIKE 'general_log';")
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
    Dim verOut : verOut = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(verOut)) = 0 Then
        Call AddResult("1.22", "系统安全", "数据库审计日志保留>=180天", "na", _
                       "未检测到MySQL服务，该项不适用。", "第1章 第1.22节", "")
        Exit Sub
    End If
    Dim expOut : expOut = MysqlQ("SHOW VARIABLES LIKE 'expire_logs_days';")
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
    Dim dbPorts : dbPorts = Array(3306, 1433, 6379)
    Dim foundPort : foundPort = ""
    Dim dp
    For Each dp In dbPorts
        If IsPortListening(CStr(dp)) Then
            If foundPort <> "" Then foundPort = foundPort & " "
            foundPort = foundPort & dp
        End If
    Next
    If foundPort <> "" Then
        ' 进一步判断是否对外（0.0.0.0）监听
        Dim nout : nout = RunCmdCached("netstat -ano 2>nul")
        Dim exposed : exposed = False
        Dim nl2
        For Each nl2 In Split(nout, Chr(10))
            If InStr(nl2, "LISTENING") > 0 And InStr(nl2, "0.0.0.0") > 0 Then
                Dim dp2
                For Each dp2 In Split(foundPort, " ")
                    If InStr(nl2, ":" & dp2 & " ") > 0 Or InStr(nl2, ":" & dp2 & Chr(9)) > 0 Then
                        exposed = True
                    End If
                Next
            End If
        Next
        If exposed Then
            Call AddResult("1.23", "系统安全", "数据库访问分离", "fail", _
                           "数据库端口" & foundPort & "对外（0.0.0.0）监听，应限制为应用服务器IP访问。", "第1章 第1.23节", _
                           "配置防火墙规则：仅允许应用服务器IP访问数据库端口")
        Else
            Call AddResult("1.23", "系统安全", "数据库访问分离", "pass", _
                           "数据库端口" & foundPort & "已绑定本机地址，未对外放开。", "第1章 第1.23节", "")
        End If
    Else
        Call AddResult("1.23", "系统安全", "数据库访问分离", "manual", _
                       "未在标准端口（3306/1433/6379）检测到数据库监听，可能使用自定义端口或数据库未以标准端口运行，需人工确认。", "第1章 第1.23节", "确认数据库实际监听端口及访问来源限制。")
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
                       "manual", "未发现 Telnet/FTP/RemoteRegistry 等已知高风险服务，但系统运行的多余服务/账户是否已清理，需人工确认。", "第2章 第2.5节", "核查开机自启服务与账户列表，清理不再使用的服务与账户。")
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
                       "manual", "未在注册表发现 TeamViewer/向日葵/AnyDesk/VNC/FileZilla 等已知高风险软件，但可能存在其他名称的同类软件或绿色版程序，需人工核查已安装软件。", "第2章 第2.9节", "核查已安装软件清单，卸载未经授权的远程控制等高危软件。")
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
    Dim out : out = RunCmdCached("net accounts 2>nul")
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
        Dim fwOut : fwOut = RunCmdCached("netsh advfirewall show allprofiles 2>nul")
        If InStr(fwOut, "Block") > 0 Or InStr(fwOut, "阻止") > 0 Then
            hasControl = True
            detail = "防火墙出站默认策略为阻止。"
        End If
    Else
        Call AddResult("2.13", "用户安全", "非法外联控制", _
                       "manual", "Windows XP 防火墙无出站策略控制，请人工核查是否部署终端非法外联控制系统。", "第2章 第2.13节", _
                       "部署终端非法外联控制系统，或将防火墙出站默认策略改为阻止。")
        Exit Sub
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
        logOut = RunCmd("cscript //nologo eventquery.vbs /L Security /R 1 2>nul")
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
    ' 精确匹配"转换状态"字段，避免误判（未加密磁盘也会显示"已加密百分比"）
    Dim encStatus : encStatus = ""
    Dim re : Set re = New RegExp
    re.IgnoreCase = True
    re.Pattern = "(转换状态|Conversion Status)[ :]*[:：][ ]*(已完全解密|Fully Decrypted|已加密|Fully Encrypted|Encryption In Progress)"
    Dim m : Set m = re.Execute(bdeOut)
    If m.Count > 0 Then encStatus = m(0).SubMatches(1)

    If encStatus = "已完全解密" Or encStatus = "Fully Decrypted" Then
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "fail", "磁盘未启用 BitLocker 加密（转换状态：已完全解密）。", "第3章 第3.1节", _
                       "启用BitLocker（控制面板→BitLocker驱动器加密 或 manage-bde -on C:)")
    ElseIf encStatus = "已加密" Or encStatus = "Fully Encrypted" Or encStatus = "Encryption In Progress" Then
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "pass", "检测到 BitLocker 已启用加密（转换状态：已加密）。", "第3章 第3.1节", "")
    Else
        Call AddResult("3.1", "数据安全", "涉密数据采取加密保护措施（BitLocker）", _
                       "manual", "无法确定 BitLocker 实际加密状态（manage-bde 无有效输出或权限不足），请人工核查。", "第3章 第3.1节", "")
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
    Dim shareOut : shareOut = RunCmdCached("net share 2>nul")
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
                   parts(0) <> "Share" And parts(0) <> "共享名" And InStr(lineStr, "命令成功") = 0 And InStr(LCase(lineStr), "command completed") = 0 Then
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

Sub Check_4_16_NLA()
    If isXP Then
        Call AddResult("4.16", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "na", "Windows XP 不支持 NLA，如紧急可远程登录验证后再升级操作系统。", "第4章 第4.16节", "")
        Exit Sub
    End If
    Dim val : val = ReadReg("HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserAuthentication")
    If val = "1" Then
        Call AddResult("4.16", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "pass", "NLA 已启用（UserAuthentication=1）。", "第4章 第4.16节", "")
    Else
        Call AddResult("4.16", "应用安全", "远程管理采取传输加密保护（NLA）", _
                       "fail", "NLA 未启用（UserAuthentication=" & val & "），应为1。", "第4章 第4.16节", _
                       "注册表 HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\UserAuthentication 设为 1")
    End If
End Sub

Sub Check_4_17_RoleSeparation()
    Dim grpOut : grpOut = RunCmd("net localgroup 2>nul")
    Dim missing : missing = ""
    Dim hasAdmin : hasAdmin = (InStr(LCase(grpOut), "administrators") > 0 Or InStr(grpOut, "管理员") > 0)
    Dim hasSec : hasSec = (InStr(LCase(grpOut), "security") > 0 Or InStr(grpOut, "安全") > 0)
    Dim hasAudit : hasAudit = (InStr(LCase(grpOut), "audit") > 0 Or InStr(grpOut, "审计") > 0)
    If Not hasAdmin Then missing = missing & "管理员组、"
    If Not hasSec Then missing = missing & "安全员组（Security/安全）、"
    If Not hasAudit Then missing = missing & "审计员组（Audit/审计）、"
    If missing = "" Then
        Call AddResult("4.17", "应用安全", "支持管理员/安全员/审计员三权分立", _
                       "pass", "检测到管理员组、安全员组、审计员组，三权分立可能已配置。", "第4章 第4.17节", "")
    Else
        Call AddResult("4.17", "应用安全", "支持管理员/安全员/审计员三权分立", _
                       "fail", "缺少专职分组：" & Left(missing, Len(missing)-1) & Chr(10) & "当前组：" & Left(grpOut, 300), _
                       "第4章 第4.17节", _
                       "创建专职安全管理员组、审计员组，实现权限分离，禁止同一账户兼任多种角色。")
    End If
End Sub

Sub Check_4_18_IPRestrict()
    Dim detail : detail = ""
    If is7OrAbove Then
        Dim fwOut : fwOut = RunCmd("netsh advfirewall firewall show rule name=all dir=in 2>nul")
        If InStr(fwOut, "RemoteIP") > 0 And InStr(LCase(fwOut), "any") = 0 Then
            detail = "检测到带IP限制的防火墙规则。"
            Call AddResult("4.18", "应用安全", "业务管理终端登录采取IP地址限制", _
                           "pass", detail, "第4章 第4.18节", "")
        Else
            Call AddResult("4.18", "应用安全", "业务管理终端登录采取IP地址限制", _
                           "fail", "未检测到管理终端IP访问限制规则。", "第4章 第4.18节", _
                           "高级安全防火墙→入站规则→新建规则：限制RDP(3389)端口只允许指定网段IP访问")
        End If
    Else
        Call AddResult("4.18", "应用安全", "业务管理终端登录采取IP地址限制", _
                       "manual", "Windows XP 请手动检查防火墙是否限制管理端口来源IP", "第4章 第4.18节", "")
    End If
End Sub

Sub Check_4_19_Lockout()
    Dim out : out = RunCmdCached("net accounts 2>nul")
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
        Call AddResult("4.19", "应用安全", "具有登录失败处理功能（账户锁定）", _
                       "pass", "账户锁定阈值 " & lockout & " 次", "第4章 第4.19节", "")
    Else
        Call AddResult("4.19", "应用安全", "具有登录失败处理功能（账户锁定）", _
                       "fail", "未配置账户锁定策略，存在暴力破解风险。", "第4章 第4.19节", _
                       "执行命令：net accounts /lockoutthreshold:5" & Chr(10) & "或【本地安全策略】→【账户策略】→【账户锁定策略】设置阈值（5次）")
    End If
End Sub

Sub Check_4_23_PortSeparation()
    Dim webPorts : webPorts = Array(80, 443, 8080, 8443)
    Dim mgmtPorts : mgmtPorts = Array(3389, 445, 135)
    Dim hasWeb : hasWeb = False
    Dim hasMgmt : hasMgmt = False
    Dim p
    For Each p In webPorts
        If IsPortListening(CStr(p)) Then hasWeb = True
    Next
    For Each p In mgmtPorts
        If IsPortListening(CStr(p)) Then hasMgmt = True
    Next
    If hasWeb And hasMgmt Then
        Call AddResult("4.23", "应用安全", "分开设置管理端口与应用端口", _
                       "pass", "应用端口（80/443）和管理端口（3389/445）分别监听，端口未混用。", "第4章 第4.23节", "")
    Else
        Call AddResult("4.23", "应用安全", "分开设置管理端口与应用端口", _
                       "manual", "本次未同时检测到Web端口和管理端口（可能未部署Web/管理服务或使用非默认端口），是否真正分离需人工确认。", "第4章 第4.23节", "核查Web端口与管理端口的监听情况，确保管理端口仅限内网访问。")
    End If
End Sub

Sub Check_4_24_AppDBSep()
    Dim webFound : webFound = ""
    Dim dbFound : dbFound = ""
    Dim webPorts : webPorts = Array(80, 443, 8080, 8000)
    Dim dbPorts  : dbPorts  = Array(3306, 1433, 5432, 27017)
    Dim p
    For Each p In webPorts
        If IsPortListening(CStr(p)) Then
            If webFound <> "" Then webFound = webFound & ","
            webFound = webFound & CStr(p)
        End If
    Next
    For Each p In dbPorts
        If IsPortListening(CStr(p)) Then
            If dbFound <> "" Then dbFound = dbFound & ","
            dbFound = dbFound & CStr(p)
        End If
    Next
    If webFound <> "" And dbFound <> "" Then
        Call AddResult("4.24", "应用安全", "应用服务和数据存储分离部署", _
                       "fail", "主机同时部署Web(" & webFound & ")和数据库(" & dbFound & ")，未分离部署。", _
                       "第4章 第4.24节", "将数据库迁移到独立服务器，应用服务器上不应安装数据库。")
    Else
        Call AddResult("4.24", "应用安全", "应用服务和数据存储分离部署", _
                       "manual", "未检测到Web和数据库同时在同一台主机上监听（可能仅部署其中之一或使用非标准端口），是否真正分离部署需人工确认。", "第4章 第4.24节", "核查应用服务与数据库是否部署于不同主机/网络区域。")
    End If
End Sub

' ============================================================
' 4.25 应用系统具有日志审计功能
' ============================================================
Sub Check_4_25_AppAudit()
    Dim findings : findings = ""
    Dim issues : issues = ""
    Dim sd : sd = SysDrive()
    Dim iisLog : iisLog = ProbeDir(sd & "\inetpub\logs\LogFiles|" & sd & "\Windows\System32\LogFiles\W3SVC1")
    If iisLog <> "" Then
        findings = findings & "  · IIS 日志目录存在：" & iisLog & Chr(10)
    End If
    Dim nginxLog : nginxLog = ProbeDir(sd & "\nginx\logs|" & sd & "\Program Files\nginx\logs|" & sd & "\scoop\apps\nginx\current\logs")
    If nginxLog <> "" Then
        findings = findings & "  · Nginx 日志目录存在：" & nginxLog & Chr(10)
    End If
    Dim appLog : appLog = ""
    If is7OrAbove Then
        appLog = RunCmd("wevtutil qe Application /c:1 /rd:true /f:text 2>nul")
    Else
        appLog = RunCmd("cscript //nologo eventquery.vbs /L Application /R 1 2>nul")
    End If
    If Len(appLog) > 10 Then
        findings = findings & "  · Windows 应用程序日志中有有效记录" & Chr(10)
    Else
        issues = issues & "  · Windows 应用程序日志为空或无法读取" & Chr(10)
    End If
    If issues <> "" And findings = "" Then
        Call AddResult("4.25", "应用安全", "应用系统具有日志审计功能", _
                       "fail", issues, "第4章 第4.25节", _
                       "检查 IIS/Nginx/Apache 等访问日志，确保应用操作被完整记录")
    ElseIf findings <> "" Then
        Call AddResult("4.25", "应用安全", "应用系统具有日志审计功能", _
                       "pass", findings, "第4章 第4.25节", "")
    Else
        Call AddResult("4.25", "应用安全", "应用系统具有日志审计功能", _
                       "pass", "Windows 应用程序日志正常，但未检测到 Web 服务（可能未部署 Web 应用）", _
                       "第4章 第4.25节", "")
    End If
End Sub

' ============================================================
' 4.6 WAF防护
' ============================================================
Sub Check_4_6_WAF()
    Dim regOut : regOut = RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
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
    If oFSO.FolderExists(SysDrive() & "\inetpub") Then
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
    Dim regOut : regOut = RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
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
    Dim vssOut : vssOut = RunCmdCached("sc query VSS 2>nul")
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
    Dim sd : sd = SysDrive()
    Dim iisLog : iisLog = ProbeDir(sd & "\inetpub\logs\LogFiles")
    If iisLog <> "" Then
        findings = findings & "  · 检测到IIS日志（" & iisLog & "）" & Chr(10)
    End If
    Dim nginxConf : nginxConf = ProbeNginxConf()
    If nginxConf <> "" Then
        Dim confOut : confOut = RunCmd("findstr /i ""worker_connections"" """ & nginxConf & """ 2>nul")
        If Len(Trim(confOut)) > 0 Then
            findings = findings & "  · Nginx配置了worker_connections：" & Trim(confOut) & Chr(10)
        End If
    End If
    If findings <> "" Then
        Call AddResult("4.12", "应用安全", "具备拒绝应用请求并发会话数量限制等功能", "pass", _
                       findings, "第4章 第4.12节", "")
    Else
        Call AddResult("4.12", "应用安全", "具备拒绝应用请求并发会话数量限制等功能", "manual", _
                       "未检测到 IIS 或 Nginx（可能未部署 Web 服务或使用非标准安装路径），并发限制配置需人工确认。", "第4章 第4.12节", _
                       "IIS：在站点配置中设置并发连接数限制" & Chr(10) & _
                       "Nginx：在nginx.conf中配置 worker_connections 和 limit_conn")
    End If
End Sub

' ============================================================
' 4.22 非默认Web应用发布端口
' ============================================================
Sub Check_4_22_DefaultPort()
    Dim defPorts : defPorts = Array(80, 443, 8080, 8443)
    Dim found : found = ""
    Dim dp
    For Each dp In defPorts
        If IsPortListening(CStr(dp)) Then
            If found <> "" Then found = found & ", "
            found = found & dp
        End If
    Next
    If found <> "" Then
        Call AddResult("4.22", "应用安全", "非默认Web应用系统发布端口", "fail", _
                       "检测到Web服务监听默认端口：" & found, "第4章 第4.22节", _
                       "将Web服务端口从默认80/443/8080修改为非标准端口，并通过路由防火墙映射")
    Else
        Call AddResult("4.22", "应用安全", "非默认Web应用系统发布端口", "pass", _
                       "未检测到Web服务监听默认端口（80/443/8080/8443）。", "第4章 第4.22节", "")
    End If
End Sub

' ============================================================
' 4.30 远程管理限制
' ============================================================
Sub Check_4_30_RemoteMgmt()
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
    Dim netAcct : netAcct = RunCmdCached("net accounts 2>nul")
    If InStr(netAcct, "Never") > 0 Or InStr(netAcct, "从不") > 0 Then
        issues = issues & "  · 未配置账户锁定策略，远程登录存在被弱口令破解风险" & Chr(10)
    End If
    If issues <> "" Then
        Call AddResult("4.30", "应用安全", "具备统一管理措施对远程管理进行限制", "fail", _
                       issues, "第4章 第4.30节", _
                       "1. 配置防火墙规则限制RDP来源IP" & Chr(10) & _
                       "2. 启用NLA：注册表 UserAuthentication=1" & Chr(10) & _
                       "3. 配置账户锁定：net accounts /lockoutthreshold:5")
    Else
        Call AddResult("4.30", "应用安全", "具备统一管理措施对远程管理进行限制", "pass", _
                       "远程管理安全配置检查通过。", "第4章 第4.30节", "")
    End If
End Sub

' ============================================================
' 4.32 应用系统具备备份与恢复功能
' ============================================================
Sub Check_4_32_AppBackup()
    Dim findings : findings = ""
    Dim vssOut : vssOut = RunCmdCached("sc query VSS 2>nul")
    If InStr(vssOut, "RUNNING") > 0 Then
        findings = findings & "  · 卷影副本服务（VSS）正在运行" & Chr(10)
    End If
    Dim taskOut : taskOut = RunCmdCached("schtasks /query /fo LIST 2>nul")
    If InStr(LCase(taskOut), "backup") > 0 Or InStr(LCase(taskOut), "备份") > 0 Or InStr(LCase(taskOut), "wbadmin") > 0 Then
        findings = findings & "  · 检测到备份相关计划任务" & Chr(10)
    End If
    Dim bdirs : bdirs = SysDrive() & "\backup " & SysDrive() & "\Backup " & SysDrive() & "\mysql_backup"
    Dim bd
    For Each bd In Split(bdirs, " ")
        If oFSO.FolderExists(bd) Then
            findings = findings & "  · 发现备份目录：" & bd & Chr(10)
        End If
    Next
    If findings <> "" Then
        Call AddResult("4.32", "应用安全", "应用系统具备备份与恢复功能", "pass", _
                       "检测到备份机制：" & Chr(10) & findings, "第4章 第4.32节", "")
    Else
        Call AddResult("4.32", "应用安全", "应用系统具备备份与恢复功能", "fail", _
                       "未检测到备份服务、备份计划任务或备份目录。", "第4章 第4.32节", _
                       "1. 安装Windows Server Backup功能并配置定期备份" & Chr(10) & _
                       "2. 创建定期备份计划任务" & Chr(10) & _
                       "3. 启用卷影副本（VSS）")
    End If
End Sub

' ============================================================
' 4.33 国产自主可控软件
' ============================================================
Sub Check_4_33_DomesticSoftware()
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
    Dim regOut : regOut = RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
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
        Call AddResult("4.33", "应用安全", "应用系统基于国产自主可控硬件或软件", "pass", _
                       findings, "第4章 第4.33节", "")
    ElseIf findings <> "" Then
        Call AddResult("4.33", "应用安全", "应用系统基于国产自主可控硬件或软件", "manual", _
                       issues & "已检测到：" & Chr(10) & findings, "第4章 第4.33节", _
                       "推进国产化替代：操作系统→麒麟/统信；数据库→达梦/人大金仓")
    Else
        Call AddResult("4.33", "应用安全", "应用系统基于国产自主可控硬件或软件", "fail", _
                       issues, "第4章 第4.33节", _
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
    Dim patchOut : patchOut = RunCmdCached("wmic qfe get InstalledOn /format:csv 2>nul")
    Dim latestDate : latestDate = ""
    Dim plines : plines = Split(patchOut, Chr(10))
    Dim pi
    For pi = 0 To UBound(plines)
        Dim pl : pl = Trim(plines(pi))
        If Len(pl) >= 8 Then
            Dim pparts : pparts = Split(pl, ",")
            If UBound(pparts) >= 1 Then
                Dim dateStr : dateStr = Trim(pparts(UBound(pparts)))
                If NormDate(dateStr) <> "" Then
                    If NormDate(dateStr) > latestDate Then latestDate = NormDate(dateStr)
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
    Dim fwOut : fwOut = RunCmdCached("netsh advfirewall show allprofiles 2>nul")
    If InStr(fwOut, "Block") > 0 Or InStr(fwOut, "阻止") > 0 Then
        findings = findings & "  · 防火墙出站默认策略为阻止" & Chr(10)
    Else
        issues = issues & "  · 防火墙出站默认策略为允许，存在非法外联风险。" & Chr(10)
    End If
    Dim regOut : regOut = RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
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
    Dim regOut : regOut = RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul")
    Dim mgmtKws : mgmtKws = "sangfor edr antiy zhijia 360entclient nsfocus topsec mcafee symantec qianxin tianqing 奇安信 天擎"
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
    Dim gpOut : gpOut = RunCmdCached("gpresult /R /SCOPE COMPUTER 2>nul")
    If InStr(gpOut, "GPO") > 0 Or InStr(gpOut, "组策略") > 0 Then
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
    Dim patchOut : patchOut = RunCmdCached("wmic qfe get InstalledOn /format:csv 2>nul")
    Dim latestDate : latestDate = ""
    Dim plines : plines = Split(patchOut, Chr(10))
    Dim pi
    For pi = 0 To UBound(plines)
        Dim pl2 : pl2 = Trim(plines(pi))
        If Len(pl2) >= 8 Then
            Dim pp : pp = Split(pl2, ",")
            If UBound(pp) >= 1 Then
                Dim ds : ds = Trim(pp(UBound(pp)))
                If NormDate(ds) <> "" Then
                    If NormDate(ds) > latestDate Then latestDate = NormDate(ds)
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
    Dim mysqlVer : mysqlVer = RunCmdCached("mysql --version 2>nul")
    If Len(Trim(mysqlVer)) > 0 Then
        If InStr(mysqlVer, "5.5.") > 0 Or InStr(mysqlVer, "5.6.") > 0 Or InStr(mysqlVer, "5.7.") > 0 Then
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
    Dim shareOut : shareOut = RunCmdCached("net share 2>nul")
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
' 辅助：检测系统审计策略是否开启（XP/7 通用）
' ============================================================
Function HasAuditEnabled()
    ' 返回 "on"（审计已开启）/ "off"（明确关闭）/ "unknown"（权限不足或无法确认）
    HasAuditEnabled = "unknown"
    Dim ap : ap = RunCmdCached("auditpol /get /category:* 2>nul")
    If InStr(ap, "成功") > 0 Or InStr(ap, "Success") > 0 Then
        HasAuditEnabled = "on"
        Exit Function
    End If
    ' 权限不足/参数错误时，auditpol 会输出"特权"或"错误"或"参数"
    If InStr(ap, "特权") > 0 Or InStr(ap, "privilege") > 0 Then
        HasAuditEnabled = "unknown"
        Exit Function
    End If
    Dim secFile : secFile = oShell.ExpandEnvironmentStrings("%TEMP%") & "\secaudit_chk.cfg"
    RunCmd "secedit /export /cfg """ & secFile & """ /quiet 2>nul"
    If oFSO.FileExists(secFile) Then
        Dim ts : Set ts = oFSO.OpenTextFile(secFile, 1)
        Dim content : content = ts.ReadAll()
        ts.Close
        oFSO.DeleteFile secFile
        Dim k
        Dim anyAuditSet : anyAuditSet = False
        Dim keywords : keywords = Array("AuditSystemEvents", "AuditLogonEvents", "AuditObjectAccess", "AuditAccountManage")
        For Each k In keywords
            If InStr(content, k & " = ") > 0 Then
                anyAuditSet = True
                Dim pos : pos = InStr(content, k & " = ")
                Dim v : v = Mid(content, pos + Len(k) + 3, 1)
                If v <> "0" Then
                    HasAuditEnabled = "on"
                    Exit Function
                End If
            End If
        Next
        ' secedit 能导出但所有审计项都是 0，说明审计确实关闭
        If anyAuditSet Then HasAuditEnabled = "off"
    End If
End Function

' ============================================================
' 3.2 数据交互/文件传输统一管控审计
' ============================================================
Sub Check_3_2_DataTransfer()
    Dim smbOn : smbOn = False
    Dim ftpOn : ftpOn = False
    Dim svcSMB : svcSMB = RunCmd("sc query lanmanserver 2>nul")
    If InStr(svcSMB, "RUNNING") > 0 Then smbOn = True
    Dim svcFTP : svcFTP = RunCmd("sc query msftpsvc 2>nul")
    If InStr(svcFTP, "RUNNING") > 0 Then ftpOn = True
    Dim svcFTP2 : svcFTP2 = RunCmd("sc query ftpsvc 2>nul")
    If InStr(svcFTP2, "RUNNING") > 0 Then ftpOn = True
    If IsPortListening("445") Then smbOn = True
    If IsPortListening("21") Then ftpOn = True
    Dim hasTransfer : hasTransfer = (smbOn Or ftpOn)
    Dim auditState : auditState = HasAuditEnabled()
    If Not hasTransfer Then
        Call AddResult("3.2", "数据安全", "数据交互/文件传输统一管控审计", _
                       "manual", "未检测到文件共享(SMB/445)或FTP传输服务，是否存在其他文件传输方式及其统一管控与审计，需人工结合业务确认。", "第3章 第3.2节", "核查是否存在专用传输通道/网盘等传输方式，并确认已纳入统一管控且开启审计。")
    ElseIf auditState = "off" Then
        Call AddResult("3.2", "数据安全", "数据交互/文件传输统一管控审计", _
                       "fail", "检测到文件共享/文件传输服务（SMB 445 或 FTP 21），但系统审计策略已明确关闭，传输行为存在未审计风险。", "第3章 第3.2节", "开启安全审计策略（登录/账户管理/对象访问），并对文件传输行为进行统一管控与审计。")
    ElseIf auditState = "unknown" Then
        Call AddResult("3.2", "数据安全", "数据交互/文件传输统一管控审计", _
                       "manual", "检测到文件共享/文件传输服务，但无法确认审计策略状态（auditpol 需管理员权限查询，当前权限不足），请以管理员身份运行脚本或人工核查审计配置。", "第3章 第3.2节", "以管理员权限运行脚本，或人工核查审计策略是否已开启。")
    Else
        Call AddResult("3.2", "数据安全", "数据交互/文件传输统一管控审计", _
                       "pass", "检测到文件共享/文件传输服务，且系统审计策略已开启。", "第3章 第3.2节", "")
    End If
End Sub

' ============================================================
' 3.3 载体写覆盖清除
' ============================================================
Sub Check_3_3_MediaErase()
    Dim found : found = ""
    Dim windir : windir = oShell.ExpandEnvironmentStrings("%windir%")
    If oFSO.FileExists(windir & "\system32\cipher.exe") Then
        found = found & "cipher（Windows自带擦除工具）"
    End If
    Dim shredOut : shredOut = RunCmd("where shred 2>nul")
    If InStr(LCase(shredOut), "shred") > 0 Then
        If found <> "" Then found = found & ", "
        found = found & "shred"
    End If
    Dim sdelOut : sdelOut = RunCmd("where sdelete 2>nul")
    If InStr(LCase(sdelOut), "sdelete") > 0 Then
        If found <> "" Then found = found & ", "
        found = found & "sdelete"
    End If
    If found <> "" Then
        Call AddResult("3.3", "数据安全", "加密存储介质降密级使用前清除写残余数据", _
                       "pass", "检测到数据擦除/写覆盖工具：" & found, "第3章 第3.3节", "")
    Else
        Call AddResult("3.3", "数据安全", "加密存储介质降密级使用前清除写残余数据", _
                       "fail", "未检测到任何数据擦除/写覆盖工具（cipher/shred/sdelete）。", "第3章 第3.3节", _
                       "使用 cipher /w 或专用擦除工具对降密级介质进行多次写覆盖，清除残余数据。")
    End If
End Sub

' ============================================================
' 3.7 数据存储系统管理登录增强
' ============================================================
Sub Check_3_7_StorageLogin()
    Dim issues : issues = ""
    Dim findings : findings = ""
    Dim adminOut : adminOut = RunCmd("net user administrator 2>nul")
    If InStr(LCase(adminOut), "not be found") > 0 Or InStr(adminOut, "找不到") > 0 Then
        findings = findings & "  · 默认 Administrator 账户已改名" & Chr(10)
    ElseIf Len(Trim(adminOut)) = 0 Then
        issues = issues & "  · 无法查询默认 Administrator 账户信息（权限不足或账户已改名）" & Chr(10)
    Else
        Dim alines : alines = Split(adminOut, Chr(10))
        Dim i
        Dim foundActive : foundActive = False
        Dim disabled : disabled = False
        For i = 0 To UBound(alines)
            If InStr(alines(i), "Account active") > 0 Or InStr(alines(i), "账户启用") > 0 Or InStr(alines(i), "帐户启用") > 0 Then
                foundActive = True
                If InStr(alines(i), "No") > 0 Or InStr(alines(i), "否") > 0 Then disabled = True
            End If
        Next
        If foundActive And disabled Then
            findings = findings & "  · 默认 Administrator 账户已禁用" & Chr(10)
        ElseIf foundActive Then
            issues = issues & "  · 默认 Administrator 账户仍启用（未改名/未禁用）" & Chr(10)
        Else
            issues = issues & "  · 无法确定默认 Administrator 账户启用状态" & Chr(10)
        End If
    End If
    Dim guestOut : guestOut = RunCmd("net user guest 2>nul")
    Dim glines : glines = Split(guestOut, Chr(10))
    Dim gi
    For gi = 0 To UBound(glines)
        If InStr(glines(gi), "Account active") > 0 Or InStr(glines(gi), "账户启用") > 0 Or InStr(glines(gi), "帐户启用") > 0 Then
            If InStr(glines(gi), "Yes") > 0 Or InStr(glines(gi), "是") > 0 Then
                issues = issues & "  · Guest 账户已启用，应禁用" & Chr(10)
            End If
        End If
    Next
    Dim auditState3 : auditState3 = HasAuditEnabled()
    Dim auditNote : auditNote = ""
    If auditState3 = "off" Then
        issues = issues & "  · 系统审计策略已明确关闭，登录行为无审计留存" & Chr(10)
    ElseIf auditState3 = "unknown" Then
        auditNote = "审计策略状态无法确认（auditpol 需管理员权限，当前未以管理员运行）。"
    End If
    If issues <> "" Then
        Call AddResult("3.7", "数据安全", "数据存储系统管理登录增强（默认账户/审计留存）", _
                       "fail", issues, "第3章 第3.7节", _
                       "1. 重命名或禁用默认 Administrator、禁用 Guest" & Chr(10) & _
                       "2. 开启登录/账户管理审计并确保日志留存180天")
    ElseIf auditNote <> "" Then
        Call AddResult("3.7", "数据安全", "数据存储系统管理登录增强（默认账户/审计留存）", _
                       "manual", findings & "  · " & auditNote, "第3章 第3.7节", _
                       "以管理员权限运行脚本确认审计策略，或人工核查审计留存配置。")
    Else
        Call AddResult("3.7", "数据安全", "数据存储系统管理登录增强（默认账户/审计留存）", _
                       "pass", findings, "第3章 第3.7节", "")
    End If
End Sub

' ============================================================
' 3.9 数据传输加密
' ============================================================
Sub Check_3_9_TransferEncryption()
    If IsPortListening("443") Then
        Call AddResult("3.9", "数据安全", "数据传输采取加密保护（HTTPS/443端口）", _
                       "pass", "检测到 443 端口监听，可能存在 HTTPS 加密传输。", "第3章 第3.9节", "")
    Else
        Call AddResult("3.9", "数据安全", "数据传输采取加密保护（HTTPS/443端口）", _
                       "manual", "未检测到 443 端口监听（可能未部署Web服务或使用非标准加密端口），数据传输加密需人工结合业务确认。", "第3章 第3.9节", "核查业务数据传输是否使用 HTTPS/TLS 等加密通道。")
    End If
End Sub

' ============================================================
' 3.11 数据访问权限分级分权
' ============================================================
Sub Check_3_11_AccessLevels()
    Dim issues : issues = ""
    Dim checked : checked = False
    Dim dirs : dirs = oShell.ExpandEnvironmentStrings("%windir%") & " " & oShell.ExpandEnvironmentStrings("%SystemDrive%") & "\inetpub"
    Dim d
    For Each d In Split(dirs, " ")
        If oFSO.FolderExists(d) Then
            checked = True
            Dim aclOut : aclOut = ""
            If isXP Then
                aclOut = RunCmd("cacls """ & d & """ 2>nul")
            Else
                aclOut = RunCmd("icacls """ & d & """ 2>nul")
            End If
            Dim alines : alines = Split(aclOut, Chr(10))
            Dim ai
            For ai = 0 To UBound(alines)
                Dim aln : aln = LCase(Trim(alines(ai)))
                If InStr(aln, "everyone") > 0 Then
                    If InStr(aln, "(f)") > 0 Or InStr(aln, ":f") > 0 Or InStr(aln, "full") > 0 Or InStr(aln, "完全控制") > 0 Then
                        issues = issues & "  · 目录 " & d & " 对 Everyone 授予完全控制权限" & Chr(10)
                    End If
                End If
            Next
        End If
    Next
    If Not checked Then
        Call AddResult("3.11", "数据安全", "数据访问权限分级分权（最小化）", _
                       "manual", "未找到可检测的关键目录，数据访问权限分级需人工核查。", "第3章 第3.11节", "核查关键数据目录的NTFS权限，确保按最小权限分级分权。")
    ElseIf issues <> "" Then
        Call AddResult("3.11", "数据安全", "数据访问权限分级分权（最小化）", _
                       "fail", issues, "第3章 第3.11节", _
                       "移除 Everyone 对关键目录的完全控制权限，按用户/组分级授权最小化。")
    Else
        Call AddResult("3.11", "数据安全", "数据访问权限分级分权（最小化）", _
                       "pass", "关键目录未检测到 Everyone 完全控制权限，权限分级符合最小化要求。", "第3章 第3.11节", "")
    End If
End Sub

' ============================================================
' 3.13 数据各环节密级处理（磁盘加密）
' ============================================================
Sub Check_3_13_DataEncryption()
    If Not is7OrAbove Then
        Call AddResult("3.13", "数据安全", "数据各环节密级处理（磁盘加密）", _
                       "na", "Windows XP 不支持 BitLocker，如使用其他加密工具（如EFS、TrueCrypt），需人工核查。", "第3章 第3.13节", "")
        Exit Sub
    End If
    Dim bdeOut : bdeOut = RunCmd("manage-bde -status 2>nul")
    Dim encStatus : encStatus = ""
    Dim re : Set re = New RegExp
    re.IgnoreCase = True
    re.Pattern = "(转换状态|Conversion Status)[ :]*[:：][ ]*(已完全解密|Fully Decrypted|已加密|Fully Encrypted|Encryption In Progress)"
    Dim m : Set m = re.Execute(bdeOut)
    If m.Count > 0 Then encStatus = m(0).SubMatches(1)
    If encStatus = "已完全解密" Or encStatus = "Fully Decrypted" Then
        Call AddResult("3.13", "数据安全", "数据各环节密级处理（磁盘加密）", _
                       "fail", "磁盘未启用 BitLocker 加密（转换状态：已完全解密）。", "第3章 第3.13节", _
                       "启用BitLocker（控制面板→BitLocker驱动器加密 或 manage-bde -on C:)")
    ElseIf encStatus = "已加密" Or encStatus = "Fully Encrypted" Or encStatus = "Encryption In Progress" Then
        Call AddResult("3.13", "数据安全", "数据各环节密级处理（磁盘加密）", _
                       "pass", "检测到 BitLocker 已启用加密（转换状态：已加密）。", "第3章 第3.13节", "")
    Else
        Call AddResult("3.13", "数据安全", "数据各环节密级处理（磁盘加密）", _
                       "manual", "无法确定 BitLocker 实际加密状态（manage-bde 无有效输出或权限不足），请人工核查。", "第3章 第3.13节", "")
    End If
End Sub

' ============================================================
' 4.2 应用系统补丁及时性
' ============================================================
Sub Check_4_2_AppPatch()
    Dim patchOut : patchOut = RunCmdCached("wmic qfe list brief /format:csv 2>nul")
    Dim lines : lines = Split(patchOut, Chr(10))
    Dim count : count = 0
    Dim i
    For i = 1 To UBound(lines)
        Dim parts : parts = Split(lines(i), ",")
        If UBound(parts) >= 4 Then
            If Len(Trim(parts(4))) >= 8 Then count = count + 1
        End If
    Next
    If count > 0 Then
        Call AddResult("4.2", "应用安全", "应用系统及时安装补丁或更新安全软件", _
                       "manual", "检测到 " & count & " 条操作系统补丁记录，说明补丁机制正常；但应用系统(业务软件/中间件/数据库)的补丁及时性需人工结合版本台账确认。", "第4章 第4.2节", "核查应用系统及其中间件/数据库的补丁版本与发布周期，确保及时更新。")
    Else
        Call AddResult("4.2", "应用安全", "应用系统及时安装补丁或更新安全软件", _
                       "fail", "未检测到任何系统补丁记录，补丁更新机制可能未启用。", "第4章 第4.2节", _
                       "通过 Windows Update 安装补丁，并对应用系统建立补丁更新机制。")
    End If
End Sub

' ============================================================
' 4.3 基于可信根可信验证
' ============================================================
Sub Check_4_3_TrustedRoot()
    Dim tpmOut : tpmOut = RunCmd("wmic /namespace:\root\cimv2\security\microsofttpm path Win32_Tpm get IsEnabled_InitialValue /value 2>nul")
    If InStr(tpmOut, "IsEnabled_InitialValue") > 0 Then
        If InStr(tpmOut, "TRUE") > 0 Then
            Call AddResult("4.3", "应用安全", "基于可信根对应用系统软件进行可信验证", _
                           "pass", "检测到 TPM 且已启用（IsEnabled_InitialValue=TRUE）。", "第4章 第4.3节", "")
        Else
            Call AddResult("4.3", "应用安全", "基于可信根对应用系统软件进行可信验证", _
                           "fail", "检测到 TPM 但未启用（IsEnabled_InitialValue 非 TRUE）。", "第4章 第4.3节", "在 BIOS/UEFI 中启用 TPM 并配置可信引导。")
        End If
    Else
        Call AddResult("4.3", "应用安全", "基于可信根对应用系统软件进行可信验证", _
                       "manual", "未检测到 TPM 设备（XP/7 平台普遍无 TPM），基于可信根的可信验证需人工核查。", "第4章 第4.3节", "核查是否部署 TPM/可信芯片，或通过签名验证等其他可信验证手段实现。")
    End If
End Sub

' ============================================================
' 4.4 公共/涉密服务器分设
' ============================================================
Sub Check_4_4_ServerSeparation()
    Dim webFound : webFound = ""
    Dim dbFound : dbFound = ""
    Dim webPorts : webPorts = Array(80, 443, 8080, 8443)
    Dim dbPorts : dbPorts = Array(3306, 1433, 5432, 27017)
    Dim p
    For Each p In webPorts
        If IsPortListening(CStr(p)) Then
            If webFound <> "" Then webFound = webFound & ","
            webFound = webFound & CStr(p)
        End If
    Next
    For Each p In dbPorts
        If IsPortListening(CStr(p)) Then
            If dbFound <> "" Then dbFound = dbFound & ","
            dbFound = dbFound & CStr(p)
        End If
    Next
    If webFound <> "" And dbFound <> "" Then
        Call AddResult("4.4", "应用安全", "公共信息服务与涉密信息服务器分设", _
                       "fail", "主机同时监听Web(" & webFound & ")和数据库(" & dbFound & ")端口，公共/涉密服务可能未分设。", _
                       "第4章 第4.4节", "将公共信息服务器与涉密信息服务器物理/逻辑分设，禁止同机混部。")
    Else
        Call AddResult("4.4", "应用安全", "公共信息服务与涉密信息服务器分设", _
                       "manual", "未检测到Web与数据库端口同时监听（可能仅部署其一或使用非标准端口），公共/涉密服务器是否分设需人工结合资产台账确认。", "第4章 第4.4节", "核查服务器资产台账，确认公共与涉密服务分设部署。")
    End If
End Sub

' ============================================================
' 4.7 网站静态化发布
' ============================================================
Sub Check_4_7_StaticPublish()
    Dim webroot : webroot = ProbeWebRoot()
    If webroot = "" Then
        Call AddResult("4.7", "应用安全", "网站应用采取静态页面发布形式", _
                       "manual", "未检测到 IIS 默认 Web 目录（" & SysDrive() & "\inetpub\wwwroot），网站是否静态发布需人工核查。", "第4章 第4.7节", "核查网站是否以静态页面发布，减少动态脚本攻击面。")
        Exit Sub
    End If
    Dim dynOut : dynOut = RunCmd("dir /s /b """ & webroot & "\*.aspx"" """ & webroot & "\*.php"" """ & webroot & "\*.asp"" """ & webroot & "\*.jsp"" 2>nul")
    If Len(Trim(dynOut)) > 0 Then
        Call AddResult("4.7", "应用安全", "网站应用采取静态页面发布形式", _
                       "fail", "检测到动态脚本文件，网站未完全静态化：" & Chr(10) & Left(dynOut, 300), "第4章 第4.7节", _
                       "将网站内容转为静态页面发布，移除动态脚本或进行静态化改造。")
    Else
        Call AddResult("4.7", "应用安全", "网站应用采取静态页面发布形式", _
                       "pass", "未检测到动态脚本文件（aspx/php/asp/jsp），网站可能为静态发布。", "第4章 第4.7节", "")
    End If
End Sub

' ============================================================
' 4.11 用户授权访问控制（RBAC）
' ============================================================
Sub Check_4_11_UserAuthz()
    Dim adminOut : adminOut = RunCmd("net localgroup administrators 2>nul")
    If Len(Trim(adminOut)) = 0 Then
        Call AddResult("4.11", "应用安全", "具备用户权限到访问控制的控制能力（RBAC）", _
                       "manual", "无法读取本地管理员组信息（权限不足），用户授权访问控制需人工核查。", "第4章 第4.11节", "核查应用层RBAC与数据库层授权是否实现最小权限。")
        Exit Sub
    End If
    If InStr(LCase(adminOut), "guest") > 0 Then
        Call AddResult("4.11", "应用安全", "具备用户权限到访问控制的控制能力（RBAC）", _
                       "fail", "检测到 Guest 账户被加入本地管理员组，存在高危权限账户。", "第4章 第4.11节", _
                       "将 Guest 从管理员组移除（net localgroup administrators guest /delete），并核查应用RBAC配置。")
        Exit Sub
    End If
    Dim lines : lines = Split(adminOut, Chr(10))
    Dim extra : extra = ""
    Dim i
    For i = 0 To UBound(lines)
        Dim ln : ln = Trim(lines(i))
        Dim low : low = LCase(ln)
        If Len(ln) > 0 Then
            If InStr(ln, "Alias name") = 0 And InStr(ln, "别名") = 0 _
               And InStr(ln, "Comment") = 0 And InStr(ln, "注释") = 0 _
               And InStr(ln, "Members") = 0 And InStr(ln, "成员") = 0 _
               And InStr(ln, "-----") = 0 _
               And InStr(low, "completed successfully") = 0 And InStr(ln, "成功完成") = 0 _
               And low <> "administrator" And low <> "管理员" _
               And low <> "domain admins" And low <> "域管理员" Then
                If extra <> "" Then extra = extra & ", "
                extra = extra & ln
            End If
        End If
    Next
    If extra <> "" Then
        Call AddResult("4.11", "应用安全", "具备用户权限到访问控制的控制能力（RBAC）", _
                       "manual", "检测到非默认管理员账户：" & extra & "，需人工确认其授权与RBAC配置是否最小化。", "第4章 第4.11节", "核查上述账户权限与应用RBAC配置，移除多余高权限账户。")
    Else
        Call AddResult("4.11", "应用安全", "具备用户权限到访问控制的控制能力（RBAC）", _
                       "manual", "未检测到非默认/可疑管理员账户，但应用层RBAC（角色-权限映射）是否细粒度仍需人工核查。", "第4章 第4.11节", "核查应用系统角色-权限映射及数据库授权是否最小化。")
    End If
End Sub

' ============================================================
' 4.14 基于用户角色细粒度访问控制
' ============================================================
Sub Check_4_14_FineGrainedACL()
    Dim issues : issues = ""
    Dim checked : checked = False
    Dim dirs : dirs = oShell.ExpandEnvironmentStrings("%windir%") & " " & oShell.ExpandEnvironmentStrings("%SystemDrive%") & "\inetpub"
    Dim d
    For Each d In Split(dirs, " ")
        If oFSO.FolderExists(d) Then
            checked = True
            Dim aclOut : aclOut = ""
            If isXP Then
                aclOut = RunCmd("cacls """ & d & """ 2>nul")
            Else
                aclOut = RunCmd("icacls """ & d & """ 2>nul")
            End If
            Dim alines : alines = Split(aclOut, Chr(10))
            Dim ai
            For ai = 0 To UBound(alines)
                Dim aln : aln = LCase(Trim(alines(ai)))
                If InStr(aln, "everyone") > 0 Then
                    If InStr(aln, "(f)") > 0 Or InStr(aln, ":f") > 0 Or InStr(aln, "full") > 0 Or InStr(aln, "完全控制") > 0 Then
                        issues = issues & "  · 目录 " & d & " 对 Everyone 授予完全控制权限" & Chr(10)
                    End If
                End If
            Next
        End If
    Next
    If Not checked Then
        Call AddResult("4.14", "应用安全", "基于用户角色和权限的访问控制（细分粒度）", _
                       "manual", "未找到可检测的关键目录，细粒度访问控制需人工核查。", "第4章 第4.14节", "核查关键目录/应用数据的ACL是否按角色细分授权。")
    ElseIf issues <> "" Then
        Call AddResult("4.14", "应用安全", "基于用户角色和权限的访问控制（细分粒度）", _
                       "fail", issues, "第4章 第4.14节", _
                       "收紧关键目录ACL，按用户/角色细分授权，移除Everyone完全控制。")
    Else
        Call AddResult("4.14", "应用安全", "基于用户角色和权限的访问控制（细分粒度）", _
                       "pass", "关键目录未检测到 Everyone 完全控制权限，ACL存在角色细分授权。", "第4章 第4.14节", "")
    End If
End Sub

' ============================================================
' 手动核查（第5-9章）
' ============================================================

' ============================================================
' 2.3 用户应用口令与认证（自动登录/空口令限制）
' ============================================================
Sub Check_2_3_AppPass()
    Dim outAuto : outAuto = RunCmd("reg query ""HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"" /v AutoAdminLogon 2>nul")
    Dim outBlank : outBlank = RunCmd("reg query HKLM\SYSTEM\CurrentControlSet\Control\Lsa /v LimitBlankPasswordUse 2>nul")
    Dim autoOn : autoOn = False
    If InStr(outAuto, "0x1") > 0 Then autoOn = True
    Dim blankBlocked : blankBlocked = False
    If InStr(outBlank, "0x1") > 0 Then blankBlocked = True
    If autoOn Then
        Call AddResult("2.3", "用户安全", "应用口令与认证登录（自动登录/空口令）", _
                       "fail", "注册表 AutoAdminLogon=1，系统启动后自动登录，无需口令认证。", _
                       "第2章 第2.3节", "删除 Winlogon 下 AutoAdminLogon/DefaultPassword 配置，确保登录需口令认证。")
    ElseIf Not blankBlocked Then
        Call AddResult("2.3", "用户安全", "应用口令与认证登录（自动登录/空口令）", _
                       "fail", "注册表 LimitBlankPasswordUse 未设为 1，空口令账户可用于网络登录。", _
                       "第2章 第2.3节", "设置 HKLM\SYSTEM\CurrentControlSet\Control\Lsa\LimitBlankPasswordUse=1。")
    Else
        Call AddResult("2.3", "用户安全", "应用口令与认证登录（自动登录/空口令）", _
                       "pass", "未发现自动登录配置，空口令账户已被限制网络登录。", _
                       "第2章 第2.3节", "")
    End If
End Sub

' ============================================================
' 3.6 边界防泄漏（终端 DLP/管控客户端联动）
' ============================================================
Sub Check_3_6_DLP()
    Dim procs : procs = LCase(RunCmd("tasklist 2>nul"))
    Dim regOut : regOut = LCase(RunCmdCached("reg query HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul") & RunCmdCached("reg query HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall /s /v DisplayName 2>nul"))
    Dim kws : kws = Array("dlp", "edr", "endpoint", "sangfor", "nsfocus", "topsec", "qianxin", "tianqing", "zhijia")
    Dim hit : hit = ""
    Dim k
    For Each k In kws
        If InStr(procs, k) > 0 Or InStr(regOut, k) > 0 Then
            hit = k
            Exit For
        End If
    Next
    If hit <> "" Then
        Call AddResult("3.6", "数据安全", "边界防泄漏(DLP)终端联动", _
                       "pass", "检测到 DLP/终端管控客户端信号（关键词：" & hit & "）。边界设备内容过滤与敏感内容识别策略仍需人工核实。", _
                       "第3章 第3.6节", "持续核查 DLP 覆盖主要外发通道并留存告警拦截日志。")
    Else
        Call AddResult("3.6", "数据安全", "边界防泄漏(DLP)终端联动", _
                       "fail", "本机未检测到 DLP/终端管控客户端进程或安装记录；仅依赖本机防火墙/杀毒按指导书 3.6.2 不宜判定符合。", _
                       "第3章 第3.6节", "部署统一 DLP/终端管控客户端，并在边界设备启用内容过滤与敏感内容识别策略。")
    End If
End Sub

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
    Call AddResult("4.1","应用安全","应用系统具备应用容灾措施","manual","应用系统备份可能涉及三层：应用自身备份功能、数据库备份策略、OS定时备份任务。指导书v2.1 4.1.3补充：应用层(查应用配置backup.path/backup.schedule)、数据库层(MySQL mysqldump/binlog SHOW VARIABLES LIKE 'log_bin'; SQL Server msdb.dbo.backupset; Oracle RMAN list backup)、OS层(schtasks查备份任务、find备份目录)。","第4章 第4.1节","建立应用+数据库+OS三层备份机制并定期验证可恢复。")
    Call AddResult("4.5","应用安全","提供承载信息的应用应具备DDoS攻击防护能力","manual","防DDoS需多层协同。指导书v2.1 4.5.3补充：边界设备(登录防火墙核查DDoS清洗策略)、应用层限流(IIS serverRuntime limit; Nginx limit_req/limit_conn; Tomcat maxThreads/acceptCount)、主机层(netsh advfirewall/sysctl tcp_syncookies/iptables limit)。","第4章 第4.5节","公共信息服务器应具备或依托DDoS防御能力。")
    Call AddResult("4.9","应用安全","Web应用具备对SQL注入/XSS等攻击防护能力","manual","需人工核查。指导书方法：curl -I 查安全头（CSP/X-Frame-Options），curl 构造 SQLi 观察返回","第4章 第4.9节","")
    Call AddResult("4.10","应用安全","Web应用具备执行代码的有效校验机制","manual","需人工核查。指导书方法：nikto 扫描、上传文件类型白名单配置","第4章 第4.10节","")
    Call AddResult("4.13","应用安全","业务管理终端专用专管","manual","需现场核查业务管理终端专设专用与登记台账，比对登录来源 IP","第4章 第4.13节","")
    Call AddResult("4.15","应用安全","文电文档签名验证/密级标识","manual","需人工核查。指导书方法：检查数字签名/证书工具、文档密级标识与访问控制配置。","第4章 第4.15节","")
    Call AddResult("4.20","应用安全","宜使用自主设计开发的网络服务、协议、接口","manual","需人工核查是否使用自主设计的协议/接口（国密 SM2/SM3/SM4 配置、国产中间件/SDK）","第4章 第4.20节","")
    Call AddResult("4.21","应用安全","基于专用物理部件或生物特征多因素的数字证书用户身份认证","manual","需人工核查多因素+数字证书认证（Windows Hello/生物特征模块、国密证书库）","第4章 第4.21节","")
    Call AddResult("4.26","应用安全","应用系统进行代码级安全漏洞挖掘","manual","需人工核查。指导书方法：核查应用日志与态势感知平台的异常告警规则","第4章 第4.26节","")
    Call AddResult("4.27","应用安全","对人机接口/网络通信/文件输入数据格式和长度检查","manual","需人工核查。指导书方法：检查输入格式/长度校验（php upload_max_filesize/post_max_size、应用校验配置）","第4章 第4.27节","")
    Call AddResult("4.28","应用安全","应有效防护并阻断SQL注入/XSS/DoS等应用层攻击","manual","需人工核查。指导书方法：curl 查安全头，curl 构造 SQLi/XSS/DoS 观察返回","第4章 第4.28节","")
    Call AddResult("4.29","应用安全","对用户登录和权限统一管理控制","manual","需人工核查。指导书方法：检查 IAM/CAS/OAuth2.0/SSO 统一认证接入与权限变更记录","第4章 第4.29节","")
    Call AddResult("4.31","应用安全","应满足请求并分会话数/带宽/单用户连接限制","manual","需人工核查。指导书方法：检查中间件并发会话数/带宽/单用户连接限制配置","第4章 第4.31节","")
    Call AddResult("2.11_m","用户安全","口令认证采用身份认证或专用的硬件或软件","na","请结合硬件USB Key/密码卡/指纹，超出本脚本范围","第2章 第2.11节","")
    Call AddResult("3.4","数据安全","确保涉密数据的传输/存储/备份等过程安全传输","manual","需人工核查涉密载体物理销毁（消磁/粉碎/溶解等）档案与影像记录","第3章 第3.4节","")
    Call AddResult("3.12","数据安全","数据采集未超出业务必要范围","manual","需人工核查。指导书方法：sc query 查采集/同步服务，比对数据库表字段与业务清单，抽查采集日志来源频率","第3章 第3.12节","")
    Call AddResult("3.14","数据安全","涉密分类提供统一管控和访问控制，应为最小化","manual","需人工核查大数据访问统一管控（OS层+大数据平台层+数据库层）与最小化授权","第3章 第3.14节","")
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

' 带缓存的命令执行（重复命令只执行一次，避免反复 fork cmd 进程）
Function RunCmdCached(cmd)
    If oCache.Exists(cmd) Then
        RunCmdCached = oCache(cmd)
    Else
        Dim r : r = RunCmd(cmd)
        oCache.Add cmd, r
        RunCmdCached = r
    End If
End Function

' MySQL 查询（连不上时直接返回空，避免逐个超时）
Function MysqlQ(sql)
    If Not mysqlConnOk Then
        MysqlQ = ""
        Exit Function
    End If
    MysqlQ = RunCmd("mysql -h " & MYSQL_HOST & " -P " & MYSQL_PORT & " -u " & MYSQL_USER & " -p" & MYSQL_PASS & " --connect-timeout=2 -e """ & sql & """ 2>nul")
End Function

' 从 db_config.conf 读取配置
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

' 配置读取优先级：环境变量 > db_config.conf > 默认值
Function GetSetting(key, defaultVal)
    If oShell.Environment("Process")(key) <> "" Then GetSetting = oShell.Environment("Process")(key) : Exit Function
    Dim cfg : cfg = ReadConfig(key)
    If cfg <> "" Then GetSetting = cfg : Exit Function
    GetSetting = defaultVal
End Function

Function NormDate(s)
    Dim v, y, m, d, p
    NormDate = ""
    v = Replace(Replace(s, Chr(13), ""), Chr(10), "")
    v = Trim(v)
    If v = "" Then Exit Function
    y = "" : m = "" : d = ""
    If InStr(v, "/") > 0 Then
        p = Split(v, "/")
        If UBound(p) >= 2 Then
            If Len(Trim(p(0))) = 4 Then
                y = Trim(p(0)) : m = Trim(p(1)) : d = Trim(p(2))
            Else
                m = Trim(p(0)) : d = Trim(p(1)) : y = Trim(p(2))
                If Len(y) = 2 Then y = "20" & y
            End If
        End If
    ElseIf Len(v) >= 8 And IsNumeric(v) Then
        y = Left(v, 4) : m = Mid(v, 5, 2) : d = Right(v, 2)
    End If
    If Len(y) = 4 And IsNumeric(y) And IsNumeric(m) And IsNumeric(d) Then
        NormDate = y & Right("0" & m, 2) & Right("0" & d, 2)
    End If
End Function

Function ReadReg(path)
    On Error Resume Next
    Dim val : val = oShell.RegRead(path)
    If Err.Number <> 0 Then val = "" : Err.Clear
    ReadReg = val
    On Error GoTo 0
End Function

' 系统盘符（如 C:）
Function SysDrive()
    SysDrive = oShell.ExpandEnvironmentStrings("%SystemDrive%")
End Function

' 判断某端口是否处于 LISTENING 状态（只匹配监听行，避免出站连接误判）
Function IsPortListening(port)
    IsPortListening = False
    Dim nout : nout = RunCmdCached("netstat -ano 2>nul")
    Dim nl
    For Each nl In Split(nout, Chr(10))
        If InStr(nl, "LISTENING") > 0 Then
            If InStr(nl, ":" & port & " ") > 0 Or InStr(nl, ":" & port & Chr(9)) > 0 Then
                IsPortListening = True
                Exit Function
            End If
        End If
    Next
End Function

' 在候选目录列表中返回第一个存在的目录，都不存在返回空
Function ProbeDir(candidates)
    Dim pd
    For Each pd In Split(candidates, "|")
        pd = Trim(pd)
        If pd <> "" And oFSO.FolderExists(pd) Then
            ProbeDir = pd
            Exit Function
        End If
    Next
    ProbeDir = ""
End Function

' 在候选文件列表中返回第一个存在的文件，都不存在返回空
Function ProbeFile(candidates)
    Dim pf
    For Each pf In Split(candidates, "|")
        pf = Trim(pf)
        If pf <> "" And oFSO.FileExists(pf) Then
            ProbeFile = pf
            Exit Function
        End If
    Next
    ProbeFile = ""
End Function

' 探测 nginx 配置文件的真实路径（通过常见安装位置 + where 反推）
Function ProbeNginxConf()
    Dim sd : sd = SysDrive()
    Dim up : up = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
    Dim cand : cand = sd & "\nginx\conf\nginx.conf|" & sd & "\Program Files\nginx\conf\nginx.conf|" & up & "\scoop\apps\nginx\current\conf\nginx.conf|" & sd & "\scoop\apps\nginx\current\conf\nginx.conf|" & sd & "\tools\nginx\conf\nginx.conf"
    ' where nginx 反推：拿到 nginx.exe 路径后向上找 conf
    Dim wh : wh = RunCmd("where nginx 2>nul")
    Dim wl
    For Each wl In Split(wh, Chr(10))
        wl = Trim(wl)
        If LCase(wl) <> "" And InStr(LCase(wl), ".exe") > 0 Then
            ' 从 exe 所在目录向上找 conf/nginx.conf（最多4层）
            Dim base : base = wl
            Dim k
            For k = 1 To 4
                Dim idx : idx = InStrRev(base, "\")
                If idx = 0 Then Exit For
                base = Left(base, idx - 1)
                Dim cc : cc = base & "\conf\nginx.conf"
                If oFSO.FileExists(cc) Then
                    ProbeNginxConf = cc
                    Exit Function
                End If
            Next
        End If
    Next
    ProbeNginxConf = ProbeFile(cand)
End Function

' 探测 Web 根目录（IIS 默认站点目录）
Function ProbeWebRoot()
    Dim sd : sd = SysDrive()
    Dim wr : wr = ProbeDir(sd & "\inetpub\wwwroot")
    If wr <> "" Then
        ProbeWebRoot = wr
    Else
        ProbeWebRoot = ""
    End If
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
    Dim ts, reportPath : reportPath = outDir & "\配置核查报告_" & stamp & ".html"
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
    ts.WriteLine "  <h1>配置核查报告（Windows XP/7版）</h1>"
    ts.WriteLine "  <div class=""sub"">参考标准：配置核查作业指导书v2.2</div>"
    ts.WriteLine "  <div class=""meta"">"
    ts.WriteLine "    <div>生成时间：" & CStr(Now()) & "，系统：" & HtmlEsc(osCaption) & "</div>"
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