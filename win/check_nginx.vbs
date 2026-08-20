' ============================================================
' 配置核查工具 - Nginx 中间件版（Windows/VBScript）
' 适用系统：Windows XP SP3 / Windows 7（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo check_nginx.vbs
' 输出文件：output\配置核查报告_Nginx_日期时间.html / .xlsx / .xls
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

Dim rID(60), rCat(60), rTitle(60), rStatus(60)
Dim rDetail(60), rChapter(60), rRec(60)
Dim rCount : rCount = 0

Dim NGINX_BIN : NGINX_BIN = ""
Dim NGINX_CONF : NGINX_CONF = ""
Dim NGINX_VERSION : NGINX_VERSION = ""

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
' 查找 nginx 与配置文件
' ============================================================
Function FindNginxBin()
    Dim r : r = RunCmd("where nginx 2>nul")
    If Trim(r) <> "" Then
        FindNginxBin = Trim(Split(r, Chr(10))(0))
        Exit Function
    End If
    FindNginxBin = ""
End Function

Function FindNginxConf()
    Dim candidates
    candidates = Array(oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\scoop\apps\nginx\current\conf\nginx.conf", "C:\nginx\conf\nginx.conf", "C:\Program Files\nginx\conf\nginx.conf")
    Dim i
    For i = 0 To UBound(candidates)
        If oFSO.FileExists(candidates(i)) Then
            FindNginxConf = candidates(i)
            Exit Function
        End If
    Next
    FindNginxConf = ""
End Function

Function ReadFile(path)
    If Not oFSO.FileExists(path) Then
        ReadFile = ""
        Exit Function
    End If
    Dim f : Set f = oFSO.OpenTextFile(path, 1, False, 0)
    ReadFile = f.ReadAll()
    f.Close
End Function

Function GetNginxVersion()
    Dim v : v = RunCmd("nginx -v 2>&1")
    Dim re : Set re = New RegExp
    re.Pattern = "nginx/([0-9.]+)"
    Dim m : Set m = re.Execute(v)
    If m.Count > 0 Then GetNginxVersion = "nginx/" & m(0).SubMatches(0)
End Function

' ============================================================
' 核查项（共8项）
' ============================================================
Sub Check_1_1()
    Dim v : v = NGINX_VERSION
    Dim mm : mm = v
    If InStr(mm, "/") > 0 Then mm = Mid(mm, InStr(mm, "/") + 1)
    Dim parts : parts = Split(mm, ".")
    Dim maj : maj = 0
    Dim min : min = 0
    If UBound(parts) >= 0 Then If IsNumeric(parts(0)) Then maj = CInt(parts(0))
    If UBound(parts) >= 1 Then If IsNumeric(parts(1)) Then min = CInt(parts(1))
    If maj < 1 Or (maj = 1 And min < 24) Then
        AddResult "1.1", "系统安全", "中间件补丁程序", "fail", "Nginx 当前版本 " & v & " 过旧（<1.24）。", "第1章", "升级到 Nginx 1.24/1.26 稳定版。"
    Else
        AddResult "1.1", "系统安全", "中间件补丁程序", "manual", "Nginx 当前版本 " & v & "，请人工核对最新补丁。", "第1章", "及时升级到最新稳定版本。"
    End If
End Sub

Sub Check_1_14()
    Dim conf : conf = NGINX_CONF
    Dim issues : issues = ""
    If conf <> "" Then
        Dim content : content = ReadFile(conf)
        Dim re
        Set re = New RegExp : re.Pattern = "server_tokens\s+off"
        If Not re.Test(content) Then issues = issues & "未隐藏版本号(server_tokens off)；"
        Set re = New RegExp : re.Pattern = "autoindex\s+on"
        If re.Test(content) Then issues = issues & "开启目录浏览(autoindex on)；"
        Set re = New RegExp : re.Pattern = "listen\s+443\s+ssl|ssl_certificate\s"
        If Not re.Test(content) Then issues = issues & "未配置HTTPS；"
    Else
        issues = "未找到 nginx 配置文件；"
    End If
    If issues = "" Then
        AddResult "1.14", "系统安全", "中间件安全加固", "pass", "Nginx 已做基础安全加固（隐藏版本号、配置 HTTPS、关闭目录浏览）。", "第1章", "持续保持中间件安全配置。"
    Else
        AddResult "1.14", "系统安全", "中间件安全加固", "fail", "Nginx 存在安全加固缺失：" & issues, "第1章", "按缺失项逐条加固。"
    End If
End Sub

Sub Check_4_5()
    Dim conf : conf = NGINX_CONF
    If conf <> "" Then
        Dim content : content = ReadFile(conf)
        Dim re : Set re = New RegExp : re.Pattern = "limit_req|limit_conn"
        If re.Test(content) Then
            AddResult "4.5", "应用安全", "防DDoS攻击能力", "pass", "已配置请求/连接限速（limit_req/limit_conn）。", "第4章", "结合前置防火墙/WAF 持续防护 DDoS。"
        Else
            AddResult "4.5", "应用安全", "防DDoS攻击能力", "manual", "未检测到限速配置，需结合前置防火墙/WAF，请人工核查。", "第4章", "配置限速或前置 WAF。"
        End If
    Else
        AddResult "4.5", "应用安全", "防DDoS攻击能力", "manual", "未找到配置文件，请人工核查。", "第4章", "配置限速。"
    End If
End Sub

Sub Check_4_6()
    Dim v : v = LCase(RunCmd("nginx -V 2>&1"))
    If InStr(v, "modsecurity") > 0 Or InStr(v, "naxsi") > 0 Or InStr(v, "lua") > 0 Then
        AddResult "4.6", "应用安全", "Web防护措施", "pass", "检测到 Web 防护模块。", "第4章", "持续保持 Web 防护规则更新。"
    Else
        AddResult "4.6", "应用安全", "Web防护措施", "manual", "未检测到 WAF 模块（modsecurity/naxsi/lua），需结合前置 WAF，请人工核查。", "第4章", "部署 WAF 或前置安全设备。"
    End If
End Sub

Sub Check_4_8()
    Dim aide : aide = Trim(RunCmd("where aide 2>nul & where tripwire 2>nul"))
    If aide <> "" Then
        AddResult "4.8", "应用安全", "网页防篡改", "pass", "检测到网页防篡改/完整性工具。", "第4章", "定期校验完整性。"
    Else
        AddResult "4.8", "应用安全", "网页防篡改", "manual", "未检测到防篡改工具（AIDE/tripwire），请人工核查。", "第4章", "部署完整性校验工具。"
    End If
End Sub

Sub Check_4_9()
    AddResult "4.9", "应用安全", "防范SQL注入/XSS", "manual", "SQL 注入/跨站脚本防护属应用层能力，需结合 WAF 与代码审计人工核查。", "第4章", "应用层参数化查询 + WAF 双重防护。"
End Sub

Sub Check_4_22()
    Dim conf : conf = NGINX_CONF
    Dim has80 : has80 = False
    If conf <> "" Then
        Dim content : content = ReadFile(conf)
        Dim re : Set re = New RegExp : re.Pattern = "listen\s+80(\s|;)"
        has80 = re.Test(content)
    End If
    If has80 Then
        AddResult "4.22", "应用安全", "更改默认服务发布端口", "fail", "仍使用默认 80 端口发布服务。", "第4章", "更改默认 80 端口为非默认端口。"
    Else
        AddResult "4.22", "应用安全", "更改默认服务发布端口", "pass", "未使用默认 80 端口发布服务。", "第4章", "持续保持非默认发布端口。"
    End If
End Sub

Sub Check_4_23()
    AddResult "4.23", "应用安全", "管理端口与应用端口分离", "na", "Nginx 无独立管理端口，此条不适用。", "第4章", "管理通道与应用通道分离。"
End Sub

' ============================================================
' 主流程
' ============================================================
NGINX_BIN = FindNginxBin()
NGINX_CONF = FindNginxConf()

WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - Nginx 中间件版（Windows/VBScript）"
WScript.Echo "  参考标准：配置核查作业指导书正式版2026_4_1"
WScript.Echo "================================================================"

If NGINX_BIN = "" Then
    WScript.Echo "  [!] 未检测到 nginx，所有项将标记为不适用。"
    AddResult "1.1", "系统安全", "中间件补丁程序", "na", "未检测到 Nginx。", "第1章", "如部署 Nginx 需及时打补丁。"
    AddResult "1.14", "系统安全", "中间件安全加固", "na", "未检测到 Nginx。", "第1章", "如部署 Nginx 需做安全加固。"
    AddResult "4.5", "应用安全", "防DDoS攻击能力", "na", "未检测到 Nginx。", "第4章", "如部署 Nginx 需配置限速。"
    AddResult "4.6", "应用安全", "Web防护措施", "na", "未检测到 Nginx。", "第4章", "如部署 Nginx 需部署 WAF。"
    AddResult "4.8", "应用安全", "网页防篡改", "na", "未检测到 Nginx。", "第4章", "如部署 Nginx 需配置防篡改。"
    AddResult "4.9", "应用安全", "防范SQL注入/XSS", "na", "未检测到 Nginx。", "第4章", "如部署 Nginx 需应用层防护。"
    AddResult "4.22", "应用安全", "更改默认服务发布端口", "na", "未检测到 Nginx。", "第4章", "如部署 Nginx 需改默认端口。"
    AddResult "4.23", "应用安全", "管理端口与应用端口分离", "na", "未检测到 Nginx。", "第4章", "管理通道与应用通道分离。"
Else
    NGINX_VERSION = GetNginxVersion()
    WScript.Echo "  Nginx 版本：" & NGINX_VERSION
    Check_1_1
    Check_1_14
    Check_4_5
    Check_4_6
    Check_4_8
    Check_4_9
    Check_4_22
    Check_4_23
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
    Dim reportPath : reportPath = outDir & "\配置核查报告_Nginx_" & stamp & ".html"
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
    ts.WriteLine "<title>Nginx 配置核查报告</title>"
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
    ts.WriteLine "<h1>Nginx 中间件配置核查报告</h1>"
    ts.WriteLine "<p style=""text-align:center;color:#666"">版本：" & HtmlEsc(NGINX_VERSION) & "　核查时间：" & CStr(Now()) & "</p>"
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

    oWS.Cells(1,1).Value = "Nginx 中间件配置核查报告"
    oWS.Cells(2,1).Value = "版本：" & NGINX_VERSION
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

    Dim xlsxPath : xlsxPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Nginx_" & stamp & ".xlsx"
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
    Dim xlsPath : xlsPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Nginx_" & stamp & ".xls"
    Dim ts
    Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7""><b>Nginx 中间件配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""7"">版本：" & HtmlEsc(NGINX_VERSION) & "</td></tr>"
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
