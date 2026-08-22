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
WScript.Echo "  参考标准：配置核查作业指导书v2.2"
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
    Dim ts, reportPath : reportPath = outDir & "\配置核查报告_Nginx_" & stamp & ".html"
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
    ts.WriteLine "  <h1>Nginx 中间件配置核查报告</h1>"
    ts.WriteLine "  <div class=""sub"">参考标准：配置核查作业指导书v2.2</div>"
    ts.WriteLine "  <div class=""meta"">"
    ts.WriteLine "    <div>版本：" & HtmlEsc(NGINX_VERSION) & "　核查时间：" & CStr(Now()) & "</div>"
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
