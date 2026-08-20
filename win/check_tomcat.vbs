' ============================================================
' 配置核查工具 - Tomcat 中间件版（Windows/VBScript）
' 适用系统：Windows XP SP3 / Windows 7（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo check_tomcat.vbs
' 连接参数：环境变量 CATALINA_HOME（默认 scoop 安装路径）
' 输出文件：output\配置核查报告_Tomcat_日期时间.html / .xlsx / .xls
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

Dim rID(60), rCat(60), rTitle(60), rStatus(60)
Dim rDetail(60), rChapter(60), rRec(60)
Dim rCount : rCount = 0

Dim CATALINA_HOME : CATALINA_HOME = ""
Dim SERVER_XML : SERVER_XML = ""
Dim TOMCAT_VERSION : TOMCAT_VERSION = ""

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

Function ReadFile(path)
    If Not oFSO.FileExists(path) Then
        ReadFile = ""
        Exit Function
    End If
    Dim f : Set f = oFSO.OpenTextFile(path, 1, False, 0)
    ReadFile = f.ReadAll()
    f.Close
End Function

Function FindCatalinaHome()
    Dim envHome : envHome = oShell.Environment("Process")("CATALINA_HOME")
    If envHome <> "" Then FindCatalinaHome = envHome : Exit Function
    Dim scoopHome : scoopHome = oShell.ExpandEnvironmentStrings("%USERPROFILE%") & "\scoop\apps\tomcat\current"
    If oFSO.FolderExists(scoopHome) Then FindCatalinaHome = scoopHome : Exit Function
    FindCatalinaHome = ""
End Function

Function GetTomcatVersion()
    Dim cmd : cmd = "java -cp """ & CATALINA_HOME & "\lib\catalina.jar"" org.apache.catalina.util.ServerInfo 2>&1"
    Dim v : v = RunCmd(cmd)
    Dim re : Set re = New RegExp
    re.Pattern = "Apache Tomcat/([0-9.]+)"
    Dim m : Set m = re.Execute(v)
    If m.Count > 0 Then GetTomcatVersion = "Apache Tomcat/" & m(0).SubMatches(0)
End Function

' ============================================================
' 核查项（共8项）
' ============================================================
Sub Check_1_1()
    Dim v : v = TOMCAT_VERSION
    Dim mm : mm = v
    Dim re : Set re = New RegExp : re.Pattern = "[0-9.]+"
    Dim m : Set m = re.Execute(v)
    Dim maj : maj = 0
    If m.Count > 0 Then
        Dim parts : parts = Split(m(0).Value, ".")
        If UBound(parts) >= 0 Then If IsNumeric(parts(0)) Then maj = CInt(parts(0))
    End If
    If maj < 9 Then
        AddResult "1.1", "系统安全", "中间件补丁程序", "fail", "Tomcat 当前版本 " & v & " 已 EOL（<9.0）。", "第1章", "升级到 Tomcat 9.0/10.1/11。"
    Else
        AddResult "1.1", "系统安全", "中间件补丁程序", "manual", "Tomcat 当前版本 " & v & "，请人工核对最新补丁。", "第1章", "及时升级到最新稳定版本。"
    End If
End Sub

Sub Check_1_14()
    Dim issues : issues = ""
    ' 默认应用
    Dim webapps : webapps = CATALINA_HOME & "\webapps"
    Dim defaults : defaults = ""
    Dim apps : apps = Array("manager", "host-manager", "docs", "examples")
    Dim j
    For j = 0 To UBound(apps)
        If oFSO.FolderExists(webapps & "\" & apps(j)) Then
            If defaults <> "" Then defaults = defaults & " "
            defaults = defaults & apps(j)
        End If
    Next
    If defaults <> "" Then issues = issues & "存在默认应用(" & defaults & ")；"
    ' HTTPS
    If oFSO.FileExists(SERVER_XML) Then
        Dim content : content = ReadFile(SERVER_XML)
        Dim re : Set re = New RegExp : re.Pattern = "SSLEnabled\s*=\s*""true"""
        If Not re.Test(content) Then issues = issues & "未配置HTTPS；"
    Else
        issues = issues & "未找到 server.xml；"
    End If
    If issues = "" Then
        AddResult "1.14", "系统安全", "中间件安全加固", "pass", "Tomcat 已做基础安全加固（配置 HTTPS、移除默认应用）。", "第1章", "持续保持中间件安全配置。"
    Else
        AddResult "1.14", "系统安全", "中间件安全加固", "fail", "Tomcat 存在安全加固缺失：" & issues, "第1章", "按缺失项逐条加固。"
    End If
End Sub

Sub Check_4_5()
    AddResult "4.5", "应用安全", "防DDoS攻击能力", "manual", "Tomcat 无内置 DDoS 防护，需结合前置防火墙/WAF，请人工核查。", "第4章", "部署前置防火墙/WAF 防 DDoS。"
End Sub

Sub Check_4_6()
    AddResult "4.6", "应用安全", "Web防护措施", "manual", "Tomcat 需前置 WAF 防护，请人工核查是否已部署。", "第4章", "部署 WAF 或前置安全设备。"
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
    Dim has8080 : has8080 = False
    If oFSO.FileExists(SERVER_XML) Then
        Dim content : content = ReadFile(SERVER_XML)
        Dim reComment : Set reComment = New RegExp : reComment.Global = True : reComment.Pattern = "<!--[\s\S]*?-->"
        content = reComment.Replace(content, "")
        Dim re : Set re = New RegExp : re.Pattern = "port=""8080"""
        has8080 = re.Test(content)
    End If
    If has8080 Then
        AddResult "4.22", "应用安全", "更改默认服务发布端口", "fail", "仍使用默认 8080 端口发布服务。", "第4章", "更改默认 8080 端口。"
    Else
        AddResult "4.22", "应用安全", "更改默认服务发布端口", "pass", "未使用默认 8080 端口发布服务。", "第4章", "持续保持非默认发布端口。"
    End If
End Sub

Sub Check_4_23()
    Dim has8005 : has8005 = False
    If oFSO.FileExists(SERVER_XML) Then
        Dim content : content = ReadFile(SERVER_XML)
        Dim reComment : Set reComment = New RegExp : reComment.Global = True : reComment.Pattern = "<!--[\s\S]*?-->"
        content = reComment.Replace(content, "")
        Dim re : Set re = New RegExp : re.Pattern = "port=""8005"""
        has8005 = re.Test(content)
    End If
    If has8005 Then
        AddResult "4.23", "应用安全", "管理端口与应用端口分离", "fail", "管理端口（shutdown）仍为默认 8005，应更改。", "第4章", "更改默认 shutdown 端口。"
    Else
        AddResult "4.23", "应用安全", "管理端口与应用端口分离", "pass", "管理端口已与默认 8005 区分。", "第4章", "持续保持管理端口与应用端口分离。"
    End If
End Sub

' ============================================================
' 主流程
' ============================================================
CATALINA_HOME = FindCatalinaHome()
If CATALINA_HOME <> "" Then SERVER_XML = CATALINA_HOME & "\conf\server.xml"

WScript.Echo "================================================================"
WScript.Echo "  配置核查工具 - Tomcat 中间件版（Windows/VBScript）"
WScript.Echo "  参考标准：配置核查作业指导书正式版2026_4_1"
WScript.Echo "================================================================"

If SERVER_XML = "" Or Not oFSO.FileExists(SERVER_XML) Then
    WScript.Echo "  [!] 未检测到 Tomcat，所有项将标记为不适用。"
    AddResult "1.1", "系统安全", "中间件补丁程序", "na", "未检测到 Tomcat。", "第1章", "如部署 Tomcat 需及时打补丁。"
    AddResult "1.14", "系统安全", "中间件安全加固", "na", "未检测到 Tomcat。", "第1章", "如部署 Tomcat 需做安全加固。"
    AddResult "4.5", "应用安全", "防DDoS攻击能力", "na", "未检测到 Tomcat。", "第4章", "如部署 Tomcat 需前置防护。"
    AddResult "4.6", "应用安全", "Web防护措施", "na", "未检测到 Tomcat。", "第4章", "如部署 Tomcat 需部署 WAF。"
    AddResult "4.8", "应用安全", "网页防篡改", "na", "未检测到 Tomcat。", "第4章", "如部署 Tomcat 需配置防篡改。"
    AddResult "4.9", "应用安全", "防范SQL注入/XSS", "na", "未检测到 Tomcat。", "第4章", "如部署 Tomcat 需应用层防护。"
    AddResult "4.22", "应用安全", "更改默认服务发布端口", "na", "未检测到 Tomcat。", "第4章", "如部署 Tomcat 需改默认端口。"
    AddResult "4.23", "应用安全", "管理端口与应用端口分离", "na", "未检测到 Tomcat。", "第4章", "管理通道与应用通道分离。"
Else
    TOMCAT_VERSION = GetTomcatVersion()
    WScript.Echo "  Tomcat 版本：" & TOMCAT_VERSION
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
    Dim reportPath : reportPath = outDir & "\配置核查报告_Tomcat_" & stamp & ".html"
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
    ts.WriteLine "<title>Tomcat 配置核查报告</title>"
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
    ts.WriteLine "<h1>Tomcat 中间件配置核查报告</h1>"
    ts.WriteLine "<p style=""text-align:center;color:#666"">版本：" & HtmlEsc(TOMCAT_VERSION) & "　核查时间：" & CStr(Now()) & "</p>"
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

    oWS.Cells(1,1).Value = "Tomcat 中间件配置核查报告"
    oWS.Cells(2,1).Value = "版本：" & TOMCAT_VERSION
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

    Dim xlsxPath : xlsxPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Tomcat_" & stamp & ".xlsx"
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
    Dim xlsPath : xlsPath = oFSO.GetAbsolutePathName(outDir) & "\配置核查报告_Tomcat_" & stamp & ".xls"
    Dim ts
    Set ts = oFSO.CreateTextFile(xlsPath, True, False)
    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""7""><b>Tomcat 中间件配置核查报告</b></td></tr>"
    ts.WriteLine "<tr><td colspan=""7"">版本：" & HtmlEsc(TOMCAT_VERSION) & "</td></tr>"
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
