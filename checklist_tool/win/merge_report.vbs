' ============================================================
' 配置核查汇总合并器（Windows/VBScript）
' 适用系统：Windows XP SP3 / 7 / 10（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo merge_report.vbs
' 功能：读取 output\ 下各组件 .xls 报告，合并成一份汇总 .xls（含"组件"列）
' ============================================================
Option Explicit

Dim oShell, oFSO
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")

' 从文件名提取组件名（配置核查报告_MySQL_xxx.xls -> MySQL；纯时间戳 -> 操作系统）
Function ComponentName(filename)
    Dim base : base = filename
    base = Replace(base, "配置核查报告_", "")
    base = Replace(base, ".xls", "")
    Dim parts : parts = Split(base, "_")
    If UBound(parts) >= 1 Then
        If IsNumeric(parts(0)) Then
            ComponentName = "操作系统"
        Else
            ComponentName = parts(0)
        End If
    Else
        ComponentName = "未知"
    End If
End Function

' HTML 转义（数据行内容可能含 & < >）
Function HtmlEsc(s)
    Dim r : r = s
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, Chr(10), "")
    HtmlEsc = r
End Function

' 状态着色
Function StatusColor(status)
    Select Case status
        Case "通过":       StatusColor = "#28a745"
        Case "未通过":     StatusColor = "#dc3545"
        Case "需人工核查": StatusColor = "#e67e00"
        Case Else:         StatusColor = "#6c757d"
    End Select
End Function

' 读取一个 .xls 报告，返回数据行数；同时把数据行写入汇总文件
Function AppendRows(ts, xlsPath, comp)
    Dim f : Set f = oFSO.OpenTextFile(xlsPath, 1, False, 0)
    Dim content : content = f.ReadAll()
    f.Close

    Dim re : Set re = New RegExp
    re.Global = True
    re.Pattern = "<tr[^>]*>([\s\S]*?)</tr>"
    Dim mtr : Set mtr = re.Execute(content)

    Dim n : n = 0
    Dim i
    For i = 0 To mtr.Count - 1
        Dim trContent : trContent = mtr(i).SubMatches(0)
        Dim retd : Set retd = New RegExp
        retd.Global = True
        retd.Pattern = "<td[^>]*>([\s\S]*?)</td>"
        Dim mtd : Set mtd = retd.Execute(trContent)
        If mtd.Count = 7 Then
            ' 数据行：编号/类别/检查项/状态/详情/修复建议/章节
            Dim status : status = mtd(3).SubMatches(0)
            ts.WriteLine "<tr>"
            ts.WriteLine "<td>" & HtmlEsc(comp) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(0).SubMatches(0)) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(1).SubMatches(0)) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(2).SubMatches(0)) & "</td>"
            ts.WriteLine "<td style=""color:" & StatusColor(status) & ";font-weight:bold"">" & HtmlEsc(status) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(4).SubMatches(0)) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(5).SubMatches(0)) & "</td>"
            ts.WriteLine "<td>" & HtmlEsc(mtd(6).SubMatches(0)) & "</td>"
            ts.WriteLine "</tr>"
            n = n + 1
        End If
    Next
    AppendRows = n
End Function

Sub GenerateMerge()
    Dim outDir : outDir = "output"
    If Not oFSO.FolderExists(outDir) Then
        WScript.Echo "[跳过] 未找到 output 目录。"
        Exit Sub
    End If

    ' 收集组件 .xls 文件
    Dim folder : Set folder = oFSO.GetFolder(outDir)
    Dim files()
    ReDim files(0)
    Dim fileCount : fileCount = 0
    Dim f
    For Each f In folder.Files
        Dim ext : ext = LCase(oFSO.GetExtensionName(f.Name))
        If ext = "xls" And InStr(f.Name, "配置核查报告") > 0 And InStr(f.Name, "汇总") = 0 Then
            ReDim Preserve files(fileCount)
            files(fileCount) = f.Path
            fileCount = fileCount + 1
        End If
    Next

    If fileCount = 0 Then
        WScript.Echo "[跳过] output 下没有组件 .xls 报告可合并。"
        Exit Sub
    End If

    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim mergePath : mergePath = outDir & "\配置核查汇总报告_" & stamp & ".xls"
    Dim ts : Set ts = oFSO.CreateTextFile(mergePath, True, False)

    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff;font-weight:bold}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""8""><b>配置核查汇总报告（共 " & fileCount & " 个组件）</b></td></tr>"
    ts.WriteLine "<tr><th class=""hd"">组件</th><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">检查项</th><th class=""hd"">状态</th><th class=""hd"">详情</th><th class=""hd"">修复建议</th><th class=""hd"">章节</th></tr>"

    Dim total : total = 0
    Dim i
    For i = 0 To fileCount - 1
        Dim comp : comp = ComponentName(oFSO.GetFileName(files(i)))
        total = total + AppendRows(ts, files(i), comp)
    Next

    ts.WriteLine "</table></body></html>"
    ts.Close

    WScript.Echo "[OK] 汇总报告已生成（共 " & total & " 项）：" & mergePath
End Sub

' ============================================================
' 主流程
' ============================================================
WScript.Echo "================================================================"
WScript.Echo "  配置核查汇总合并器（Windows/VBScript）"
WScript.Echo "================================================================"
Call GenerateMerge()
