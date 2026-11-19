# ============================================================
# 零依赖 xlsx 合并器（PowerShell + .NET，无需 Python/openpyxl）
# 把 win\output 下各组件 xlsx 合并成一个汇总 xlsx
# 用法：powershell -ExecutionPolicy Bypass -File merge_xlsx.ps1 [合并目录]
#
# 两种用法（评审 2026-09-30 第 6 条：不同设备的核查报告要能合并整合）
#   1) 单设备：output\ 下直接放各组件 xlsx      -> 设备列显示「本机」
#   2) 多设备：output\服务器A\、output\服务器B\  -> 设备列填子目录名
#
# 列结构自适应：按表头文字识别，兼容
#   7 列（旧组件报告） 编号/类别/检查项/状态/详情/修复建议/章节
#   8 列（网络设备版） 章节/编号/类别/核查项/结果/详情/建议/参考指导书
#   10 列（整改后）    章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书
# 输出 13 列：设备/组件/章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书/数据目录
# ============================================================
param([string]$OutDir)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression

if (-not $OutDir) {
    $OutDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "output"
}
if (-not (Test-Path $OutDir)) { Write-Host "[跳过] 未找到目录：$OutDir"; exit 0 }

function Xml-Unescape([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&#10;', "`n"
    $s = $s -replace '&quot;', '"' -replace '&apos;', "'"
    $s = $s -replace '&lt;', '<' -replace '&gt;', '>'
    $s = $s -replace '&amp;', '&'
    return $s
}
function Xml-Esc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;' -replace "'", '&apos;'
    $s = $s -replace "`r`n", '&#10;' -replace "`n", '&#10;' -replace "`r", ''
    return $s
}

# 读取 sharedStrings.xml（经 Excel/WPS 另存过的 xlsx 用共享字符串，单元格只有 <v> 索引）
function Read-SharedStrings($zip) {
    $ss = New-Object System.Collections.ArrayList
    $e = $zip.GetEntry('xl/sharedStrings.xml')
    if ($null -eq $e) { return $ss }
    $r = New-Object System.IO.StreamReader($e.Open(), [System.Text.Encoding]::UTF8)
    $x = $r.ReadToEnd(); $r.Close()
    foreach ($m in [regex]::Matches($x, '<si>(.*?)</si>', 'Singleline')) {
        $txt = ''
        foreach ($t in [regex]::Matches($m.Groups[1].Value, '<t[^>]*>(.*?)</t>', 'Singleline')) {
            $txt += (Xml-Unescape $t.Groups[1].Value)
        }
        [void]$ss.Add($txt)
    }
    return $ss
}

# 读取 xlsx 的 sheet1.xml，返回 @{ Header = 表头数组; Rows = 数据行数组 }
# 单元格两种存法都支持：inlineStr（本工具生成）/ sharedStrings（Excel、WPS 重存）
# 按 <c r="B4"> 的列字母定位，不按出现顺序——空格子会被跳过，顺序法会串列
function Read-SheetTable([string]$xlsxFile) {
    try {
        $fs = [System.IO.File]::OpenRead($xlsxFile)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
        if (-not $entry) { $zip.Dispose(); $fs.Close(); return $null }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $xml = $reader.ReadToEnd()
        $reader.Close()
        $ss = Read-SharedStrings $zip
        $zip.Dispose(); $fs.Close()
    } catch { return $null }

    # 元信息：源报告第 3 行的「数据目录：…」（数据库类报告才有）
    $dataDir = ''
    $dm = [regex]::Match($xml, '数据目录：([^<　\r\n]*)')
    if ($dm.Success) { $dataDir = $dm.Groups[1].Value.Trim() }

    $rows = @()
    foreach ($m in [regex]::Matches($xml, '<row[^>]*>(.*?)</row>', 'Singleline')) {
        $cells = @{}
        $maxIdx = -1
        foreach ($cm in [regex]::Matches($m.Groups[1].Value, '<c\s[^>]*?(?:/>|>.*?</c>)', 'Singleline')) {
            $tag = $cm.Value
            $rm = [regex]::Match($tag, 'r="([A-Z]+)\d+"')
            if (-not $rm.Success) { continue }
            $colIdx = 0
            foreach ($ch in $rm.Groups[1].Value.ToCharArray()) { $colIdx = $colIdx * 26 + ([int][char]$ch - 64) }
            $colIdx = $colIdx - 1
            $type = ''
            $tm = [regex]::Match($tag, 't="([^"]+)"')
            if ($tm.Success) { $type = $tm.Groups[1].Value }
            $val = ''
            if ($type -eq 'inlineStr') {
                $im = [regex]::Match($tag, '<is>.*?<t[^>]*>(.*?)</t>', 'Singleline')
                if ($im.Success) { $val = Xml-Unescape $im.Groups[1].Value }
            } else {
                $vm = [regex]::Match($tag, '<v>(.*?)</v>', 'Singleline')
                if ($vm.Success) {
                    if ($type -eq 's') {
                        $idx = 0
                        if ([int]::TryParse($vm.Groups[1].Value, [ref]$idx)) {
                            if ($idx -ge 0 -and $idx -lt $ss.Count) { $val = $ss[$idx] }
                        }
                    } else {
                        # 数值单元格：Excel 存的是浮点（1.1 会写成 1.1000000000000001），
                        # 转文本时按 15 位有效数字收敛，避免汇总里出现浮点尾巴
                        $dv = 0.0
                        # xlsx 数值恒为「.」小数点的不变区域格式，不能跟随本机区域设置
                        if ([double]::TryParse($vm.Groups[1].Value,
                                [System.Globalization.NumberStyles]::Float,
                                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dv)) {
                            $val = $dv.ToString('G15', [System.Globalization.CultureInfo]::InvariantCulture)
                        } else {
                            $val = Xml-Unescape $vm.Groups[1].Value
                        }
                    }
                }
            }
            $cells[$colIdx] = $val
            if ($colIdx -gt $maxIdx) { $maxIdx = $colIdx }
        }
        if ($maxIdx -ge 0) {
            $arr = @()
            for ($ci = 0; $ci -le $maxIdx; $ci++) {
                if ($cells.ContainsKey($ci)) { $arr += $cells[$ci] } else { $arr += '' }
            }
            $rows += ,$arr
        }
    }
    if ($rows.Count -lt 2) { return $null }

    # 表头行：含「章节」或「编号」的那一行（旧表第一列是编号，新表是章节）
    $hdr = -1
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($rows[$i] -contains '章节' -or $rows[$i] -contains '编号') { $hdr = $i; break }
    }
    if ($hdr -lt 0 -or $hdr -ge ($rows.Count - 1)) { return $null }

    $data = @()
    if ($hdr + 1 -le $rows.Count - 1) { $data = @($rows[($hdr + 1)..($rows.Count - 1)]) }
    return @{ Header = $rows[$hdr]; Rows = $data; DataDir = $dataDir }
}

# 按候选列名取索引
function Get-ColIdx($header, [string[]]$names) {
    foreach ($n in $names) {
        for ($i = 0; $i -lt $header.Count; $i++) {
            if ($header[$i] -eq $n) { return $i }
        }
    }
    return -1
}

function Get-Cell($row, [int]$idx) {
    if ($idx -lt 0 -or $idx -ge $row.Count) { return '' }
    return $row[$idx]
}

# 从文件名提取组件名（配置核查报告_MySQL_2026... -> MySQL；纯时间戳 -> 操作系统）
function Get-Component([string]$name) {
    $base = $name -replace '^配置核查报告_', '' -replace '\.xlsx$', ''
    $parts = $base -split '_'
    if ($parts.Count -ge 2 -and $parts[0] -notmatch '^\d+$') { return $parts[0] }
    return '操作系统'
}

# ---------- 收集源文件（支持设备子目录） ----------
$sources = @{}   # key = "设备|组件" -> @{Device;Component;Path;Time}
$subDirs = @(Get-ChildItem -Path $OutDir -Directory -ErrorAction SilentlyContinue)

function Add-Sources([string]$dir, [string]$dev) {
    $list = @(Get-ChildItem -Path (Join-Path $dir '配置核查报告_*.xlsx') -ErrorAction SilentlyContinue |
              Where-Object { $_.Name -notlike '配置核查汇总报告*' })
    foreach ($f in $list) {
        $comp = Get-Component $f.Name
        $key = "$dev|$comp"
        if (-not $sources.ContainsKey($key) -or $f.LastWriteTime -gt $sources[$key].Time) {
            $sources[$key] = @{ Device = $dev; Component = $comp; Path = $f.FullName; Time = $f.LastWriteTime }
        }
    }
}

# 根目录的报告一律按「本机」收录；同时扫描设备子目录。
# 不能因存在子目录就丢掉根目录报告（旧版这样写会静默漏数据）。
Add-Sources $OutDir '本机'
foreach ($d in $subDirs) { Add-Sources $d.FullName $d.Name }

if ($sources.Count -eq 0) { Write-Host "[跳过] $OutDir 下没有组件 xlsx 可合并"; exit 0 }

# ---------- 逐份读取并按列名映射 ----------
$allRows = @()
$devSet = @{}
$hasNetDev = $false
foreach ($k in $sources.Keys) { if ($k -like '*网络设备*') { $hasNetDev = $true } }
foreach ($k in ($sources.Keys | Sort-Object)) {
    $src = $sources[$k]
    $devSet[$src.Device] = 1
    $t = Read-SheetTable $src.Path
    if ($null -eq $t) { Write-Host "  [跳过] 解析失败：$($src.Path)"; continue }
    $h = $t.Header
    # 「章节」与「参考指导书」是两个不同的源列，不能共用一个索引：
    # 10 列报告里 参考指导书 是第 10 列，共用会把整段书名塞进「章节」
    $iCh     = Get-ColIdx $h @('章节')
    $iRef    = Get-ColIdx $h @('参考指导书', '章节')
    $iId     = Get-ColIdx $h @('编号')
    $iCat    = Get-ColIdx $h @('类别', '分类')
    $iTitle  = Get-ColIdx $h @('核查项', '检查项')
    $iStatus = Get-ColIdx $h @('结果', '状态')
    $iMethod = Get-ColIdx $h @('验证过程/方法', '验证过程', '方法')
    $iDetail = Get-ColIdx $h @('详情')
    $iRec    = Get-ColIdx $h @('修改建议', '建议', '修复建议')
    $iReq    = Get-ColIdx $h @('核查要求')
    if ($iId -lt 0 -or $iStatus -lt 0) { Write-Host "  [跳过] 表头不含编号/结果：$($src.Path)"; continue }

    foreach ($r in $t.Rows) {
        $id = Get-Cell $r $iId
        $title = Get-Cell $r $iTitle
        if (-not $id -and -not $title) { continue }
        $chVal = Get-Cell $r $iCh
        # 与 VBS / 麒麟版一致：网络设备报告已在时，丢弃操作系统脚本第5章的 na 占位行
        if ($hasNetDev -and $src.Component -eq '操作系统' -and $chVal -like '第5章*') { continue }
        $allRows += ,@(
            $src.Device, $src.Component,
            $chVal, $id, (Get-Cell $r $iCat), $title,
            (Get-Cell $r $iStatus), (Get-Cell $r $iMethod), (Get-Cell $r $iDetail),
            (Get-Cell $r $iRec), (Get-Cell $r $iReq), (Get-Cell $r $iRef),
            $(if ($t.DataDir) { $t.DataDir } else { '-' })
        )
    }
}
if ($allRows.Count -eq 0) { Write-Host "[跳过] 未提取到数据行"; exit 0 }

# ---------- 构建 sheet XML（12 列 A-L） ----------
function Col([int]$n) { [char](64 + $n) }
$headers = @('设备', '组件', '章节', '编号', '类别', '核查项', '结果',
             '验证过程/方法', '详情', '修改建议', '核查要求', '参考指导书', '数据目录')
$nCol = $headers.Count
$sb = New-Object System.Text.StringBuilder
$title = "配置核查汇总报告（$($devSet.Count) 台设备 / $($sources.Count) 份组件报告 / $($allRows.Count) 项）"
[void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>' + (Xml-Esc $title) + '</t></is></c></row>')
[void]$sb.Append('<row r="4">')
for ($c = 0; $c -lt $nCol; $c++) {
    $ref = (Col ($c + 1)) + '4'
    [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + (Xml-Esc $headers[$c]) + '</t></is></c>')
}
[void]$sb.Append('</row>')
$ri = 0
foreach ($r in $allRows) {
    $ri++
    $rn = $ri + 4
    [void]$sb.Append('<row r="' + $rn + '">')
    for ($c = 0; $c -lt $nCol; $c++) {
        $ref = (Col ($c + 1)) + $rn
        $v = ''
        if ($c -lt $r.Count) { $v = $r[$c] }
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t xml:space="preserve">' + (Xml-Esc $v) + '</t></is></c>')
    }
    [void]$sb.Append('</row>')
}
$sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + $sb.ToString() + '</sheetData></worksheet>'

$ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + "`n" + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + "`n" + '<Default Extension="xml" ContentType="application/xml"/>' + "`n" + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + "`n" + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + "`n" + '</Types>'
$rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' + "`n" + '</Relationships>'
$wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + "`n" + '<sheets><sheet name="核查汇总" sheetId="1" r:id="rId1"/></sheets>' + "`n" + '</workbook>'
$wbrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' + "`n" + '</Relationships>'

Add-Type -AssemblyName System.IO.Compression.FileSystem
$enc = New-Object System.Text.UTF8Encoding($false)
$stamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outFile = Join-Path $OutDir "配置核查汇总报告_$stamp.xlsx"
if (Test-Path $outFile) { Remove-Item $outFile -Force }

$fs = [System.IO.File]::Create($outFile)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
function Add-ZipEntry($z, $name, $content) {
    $entry = $z.CreateEntry($name)
    $sw = New-Object System.IO.StreamWriter($entry.Open(), $enc)
    $sw.Write($content)
    $sw.Close()
}
Add-ZipEntry $zip '[Content_Types].xml' $ct
Add-ZipEntry $zip '_rels/.rels' $rels
Add-ZipEntry $zip 'xl/workbook.xml' $wb
Add-ZipEntry $zip 'xl/_rels/workbook.xml.rels' $wbrels
Add-ZipEntry $zip 'xl/worksheets/sheet1.xml' $sheetXml
$zip.Dispose()
$fs.Close()

Write-Host "汇总 Excel(.xlsx)报告已生成（$($devSet.Count) 台设备 / $($sources.Count) 份组件报告 / $($allRows.Count) 项）："
Write-Host "  $outFile"
