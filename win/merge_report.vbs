' ============================================================
' 配置核查汇总合并器（Windows/VBScript）
' 适用系统：Windows XP SP3 / 7 / 10（无需 PowerShell/Python）
' 运行方式：cscript //NoLogo merge_report.vbs [目录]
'           不带参数时默认合并 output\
' 功能：读取指定目录（及其子目录）下的组件 .xls 报告，合并成一份汇总 .xls。
'
' 两种用法
'   1) 单设备：output\ 下直接放各组件报告 -> 输出「组件」列
'   2) 多设备：output\服务器A\、output\服务器B\ 各放一套 -> 输出「设备」列
'      （多台被测机分别跑完 run_all 后，把各自 output 拷成子目录即可合并整合）
'
' 列结构自适应：按报告表头文字识别列，兼容
'   7 列（2026-09 前组件报告） 编号/类别/检查项/状态/详情/修复建议/章节
'   8 列（网络设备版）         章节/编号/类别/核查项/结果/详情/建议/参考指导书
'   10 列（2026-09-30 整改后） 章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书
' 汇总输出 14 列：设备/组件/章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书/数据目录/报告文件
' ============================================================
Option Explicit

Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

' 是否已收集到网络设备报告（决定是否丢弃单机脚本第5章的 na 占位行）
Dim HAS_NETDEV : HAS_NETDEV = False

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
        Case "通过", "合规": StatusColor = "#28a745"
        Case "未通过", "不合规": StatusColor = "#dc3545"
        Case "需人工核查": StatusColor = "#e67e00"
        Case Else:         StatusColor = "#6c757d"
    End Select
End Function

' 去 HTML 标签、压缩空白
Function StripTags(s)
    Dim re : Set re = New RegExp
    re.Global = True
    re.Pattern = "<[^>]+>"
    Dim t : t = re.Replace(s, "")
    t = Replace(t, "&nbsp;", " ")
    t = Replace(t, "&amp;", "&")
    t = Replace(t, "&lt;", "<")
    t = Replace(t, "&gt;", ">")
    t = Replace(t, "&quot;", """")
    t = Replace(t, "&#39;", "'")
    t = Replace(t, "&nbsp;", " ")
    StripTags = Trim(t)
End Function

' 表头列名 -> 索引（Dictionary）
Function BuildColMap(content)
    Dim re : Set re = New RegExp
    re.Global = True
    re.Pattern = "<th[^>]*>([\s\S]*?)</th>"
    Dim mt : Set mt = re.Execute(content)
    Dim m : Set m = CreateObject("Scripting.Dictionary")
    Dim i
    For i = 0 To mt.Count - 1
        Dim name : name = StripTags(mt(i).SubMatches(0))
        If name <> "" And Not m.Exists(name) Then m.Add name, i
    Next
    Set BuildColMap = m
End Function

' 按候选名（| 分隔）取列索引，取不到返回 -1
Function ColIdx(map, names)
    Dim a : a = Split(names, "|")
    Dim i
    ColIdx = -1
    For i = 0 To UBound(a)
        If map.Exists(a(i)) Then
            ColIdx = map(a(i))
            Exit Function
        End If
    Next
End Function

' 安全取单元格文本
Function CellVal(mtd, idx, n)
    If idx < 0 Or idx >= n Then
        CellVal = ""
    Else
        CellVal = mtd(idx).SubMatches(0)
    End If
End Function

' 读取一个 .xls 报告，返回数据行数；同时把数据行写入汇总文件
Function AppendRows(ts, xlsPath, comp, dev)
    Dim f : Set f = oFSO.OpenTextFile(xlsPath, 1, False, 0)
    Dim content : content = f.ReadAll
    f.Close

    ' 汇总表「数据目录」列（评审第7条）：从源报告 meta 行取「数据目录：…」
    Dim sData : sData = "-"
    Dim reDD : Set reDD = New RegExp
    reDD.Global = False
    reDD.Pattern = "数据目录：([^<" & Chr(10) & Chr(13) & "　]*)"
    Dim mDD : Set mDD = reDD.Execute(content)
    If mDD.Count > 0 Then sData = Trim(mDD(0).SubMatches(0))
    If sData = "" Then sData = "-"
    ' 汇总表「报告文件」列：本行来源的组件报告文件名（便于按汇总记录回查归档）
    Dim sFile : sFile = oFSO.GetFileName(xlsPath)

    Dim cm : Set cm = BuildColMap(content)
    Dim iId     : iId     = ColIdx(cm, "编号")
    Dim iCat    : iCat    = ColIdx(cm, "类别|分类")
    Dim iTitle  : iTitle  = ColIdx(cm, "核查项|检查项")
    Dim iStatus : iStatus = ColIdx(cm, "结果|状态")
    Dim iMethod : iMethod = ColIdx(cm, "验证过程/方法|验证过程|方法")
    Dim iDetail : iDetail = ColIdx(cm, "详情")
    Dim iRec    : iRec    = ColIdx(cm, "修改建议|建议|修复建议")
    Dim iReq    : iReq    = ColIdx(cm, "核查要求")
    Dim iCh     : iCh     = ColIdx(cm, "章节")
    Dim iRef    : iRef    = ColIdx(cm, "参考指导书|章节")

    If iId < 0 Or iStatus < 0 Then
        AppendRows = 0
        Exit Function
    End If

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
        Dim ntd : ntd = mtd.Count
        If ntd >= 5 Then
            Dim sId : sId = StripTags(CellVal(mtd, iId, ntd))
            Dim sTitle : sTitle = StripTags(CellVal(mtd, iTitle, ntd))
            If sId <> "" Or sTitle <> "" Then
                Dim sCh : sCh = StripTags(CellVal(mtd, iCh, ntd))
                Dim sRef : sRef = StripTags(CellVal(mtd, iRef, ntd))
                Dim skipRow : skipRow = False
                ' 第5章在核查表对象矩阵中属"网络设备"对象：网络设备报告已在时，
                ' 丢弃单机脚本的第5章 na 占位行，避免同一章两套结论
                If HAS_NETDEV And comp = "操作系统" And Left(sCh, 3) = "第5章" Then skipRow = True
                If Not skipRow Then
                    Dim status : status = StripTags(CellVal(mtd, iStatus, ntd))
                    ts.WriteLine "<tr>"
                    ts.WriteLine "<td>" & HtmlEsc(dev) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(comp) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sCh) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sId) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(StripTags(CellVal(mtd, iCat, ntd))) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sTitle) & "</td>"
                    ts.WriteLine "<td style=""color:" & StatusColor(status) & ";font-weight:bold"">" & HtmlEsc(status) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(StripTags(CellVal(mtd, iMethod, ntd))) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(StripTags(CellVal(mtd, iDetail, ntd))) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(StripTags(CellVal(mtd, iRec, ntd))) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(StripTags(CellVal(mtd, iReq, ntd))) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sRef) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sData) & "</td>"
                    ts.WriteLine "<td>" & HtmlEsc(sFile) & "</td>"
                    ts.WriteLine "</tr>"
                    n = n + 1
                End If
            End If
        End If
    Next
    AppendRows = n
End Function

' 递归收集：key = 设备 & "|" & 组件，value = 文件路径（同键保留最新）
Sub ScanFolder(dir, dev, fileMap)
    Dim folder : Set folder = oFSO.GetFolder(dir)
    Dim f
    For Each f In folder.Files
        Dim ext : ext = LCase(oFSO.GetExtensionName(f.Name))
        If ext = "xls" And InStr(f.Name, "配置核查报告") > 0 And InStr(f.Name, "汇总") = 0 Then
            Dim cname : cname = ComponentName(f.Name)
            Dim key : key = dev & "|" & cname
            If InStr(f.Name, "网络设备") > 0 Then HAS_NETDEV = True
            If fileMap.Exists(key) Then
                If f.DateLastModified > oFSO.GetFile(fileMap(key)).DateLastModified Then fileMap(key) = f.Path
            Else
                fileMap.Add key, f.Path
            End If
        End If
    Next
    Dim sf
    For Each sf In folder.SubFolders
        ScanFolder sf.Path, sf.Name, fileMap
    Next
End Sub

Sub GenerateMerge(rootDir)
    If Not oFSO.FolderExists(rootDir) Then
        WScript.Echo "[跳过] 未找到目录：" & rootDir
        Exit Sub
    End If

    Dim fileMap : Set fileMap = CreateObject("Scripting.Dictionary")
    ScanFolder rootDir, "", fileMap

    If fileMap.Count = 0 Then
        WScript.Echo "[跳过] " & rootDir & " 下没有组件 .xls 报告可合并。"
        Exit Sub
    End If

    Dim dtNow : dtNow = Now()
    Dim stamp : stamp = Year(dtNow) & Right("0" & Month(dtNow), 2) & Right("0" & Day(dtNow), 2) & _
                        "_" & Right("0" & Hour(dtNow), 2) & Right("0" & Minute(dtNow), 2) & _
                        Right("0" & Second(dtNow), 2)
    Dim mergePath : mergePath = rootDir & "\配置核查汇总报告_" & stamp & ".xls"
    Dim ts : Set ts = oFSO.CreateTextFile(mergePath, True, False)

    Dim devSet : Set devSet = CreateObject("Scripting.Dictionary")
    Dim k, devKey
    For Each k In fileMap.Keys
        devKey = Left(k, InStr(k, "|") - 1)
        If devKey = "" Then devKey = "本机"
        If Not devSet.Exists(devKey) Then devSet.Add devKey, 1
    Next
    Dim devCount : devCount = devSet.Count
    If devCount = 0 Then devCount = 1

    ts.WriteLine "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
    ts.WriteLine "<head><meta http-equiv=""Content-Type"" content=""text/html; charset=GBK"">"
    ts.WriteLine "<style>td,th{border:1px solid #bbb;padding:4px 6px}th.hd{background:#007bff;color:#fff;font-weight:bold}</style></head><body>"
    ts.WriteLine "<table cellspacing=""0"">"
    ts.WriteLine "<tr><td colspan=""14""><b>配置核查汇总报告（" & devCount & " 台设备 / " & fileMap.Count & " 份组件报告）</b></td></tr>"
    ts.WriteLine "<tr><th class=""hd"">设备</th><th class=""hd"">组件</th><th class=""hd"">章节</th><th class=""hd"">编号</th><th class=""hd"">类别</th><th class=""hd"">核查项</th><th class=""hd"">结果</th><th class=""hd"">验证过程/方法</th><th class=""hd"">详情</th><th class=""hd"">修改建议</th><th class=""hd"">核查要求</th><th class=""hd"">参考指导书</th><th class=""hd"">数据目录</th><th class=""hd"">报告文件</th></tr>"

    Dim total : total = 0
    For Each k In fileMap.Keys
        Dim pos : pos = InStr(k, "|")
        Dim dev : dev = Left(k, pos - 1)
        Dim comp : comp = Mid(k, pos + 1)
        If dev = "" Then dev = "本机"
        total = total + AppendRows(ts, fileMap(k), comp, dev)
    Next

    ts.WriteLine "</table></body></html>"
    ts.Close

    WScript.Echo "[OK] 汇总报告已生成（" & devCount & " 台设备 / " & total & " 项）：" & mergePath
End Sub

' ============================================================
' 主流程
' ============================================================
Dim argDir : argDir = "output"
If WScript.Arguments.Count >= 1 Then argDir = WScript.Arguments(0)

WScript.Echo "================================================================"
WScript.Echo "  配置核查汇总合并器（Windows/VBScript）"
WScript.Echo "  合并目录：" & argDir
WScript.Echo "  （含子目录时按设备名分列；单目录时设备列显示「本机」）"
WScript.Echo "================================================================"
Call GenerateMerge(argDir)
