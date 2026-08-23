# ============================================================
# 零依赖 xlsx 合并器（PowerShell + .NET，无需 Python/openpyxl）
# 把 win\output 下各组件 xlsx 合并成一个汇总 xlsx（单 sheet，加"组件"列）
# 用法：powershell -ExecutionPolicy Bypass -File merge_xlsx.ps1 [outDir]
# ============================================================
param([string]$OutDir)

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName System.IO.Compression

if (-not $OutDir) {
    $OutDir = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "output"
}
if (-not (Test-Path $OutDir)) { Write-Host "[跳过] 未找到 output 目录：$OutDir"; exit 0 }

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

# 读取 xlsx 的 sheet1.xml，返回数据行（8 列，跳过标题/汇总/空/表头前 4 行）
function Read-SheetData([string]$xlsxFile) {
    try {
        $fs = [System.IO.File]::OpenRead($xlsxFile)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
        $entry = $zip.GetEntry('xl/worksheets/sheet1.xml')
        if (-not $entry) { $zip.Dispose(); $fs.Close(); return @() }
        $reader = New-Object System.IO.StreamReader($entry.Open(), [System.Text.Encoding]::UTF8)
        $xml = $reader.ReadToEnd()
        $reader.Close(); $zip.Dispose(); $fs.Close()
    } catch { return @() }

    $rows = @()
    foreach ($m in [regex]::Matches($xml, '<row[^>]*>(.*?)</row>', 'Singleline')) {
        $cells = @()
        foreach ($c in [regex]::Matches($m.Groups[1].Value, '<t[^>]*>(.*?)</t>', 'Singleline')) {
            $cells += (Xml-Unescape $c.Groups[1].Value)
        }
        $rows += ,$cells
    }
    if ($rows.Count -le 3) { return @() }
    # 找到表头行（第一列"章节"），从其下一行开始取数据
    $start = 0
    for ($i = 0; $i -lt $rows.Count; $i++) {
        if ($rows[$i].Count -gt 0 -and $rows[$i][0] -eq '章节') { $start = $i + 1; break }
    }
    if ($start -ge $rows.Count) { return @() }
    return ,@($rows[$start..($rows.Count - 1)])
}

# 从文件名提取组件名（配置核查报告_MySQL_2026... -> MySQL；纯时间戳 -> 操作系统）
function Get-Component([string]$name) {
    $base = $name -replace '^配置核查报告_', '' -replace '\.xlsx$', ''
    $parts = $base -split '_'
    if ($parts.Count -ge 2 -and $parts[0] -notmatch '^\d+$') { return $parts[0] }
    return '操作系统'
}

$files = @(Get-ChildItem "$OutDir\配置核查报告_*.xlsx" | Where-Object { $_.Name -notlike '配置核查汇总报告*' } | Sort-Object Name)
if ($files.Count -eq 0) { Write-Host "[跳过] $OutDir 下没有组件 xlsx 可合并"; exit 0 }

$headers = @('组件','章节','编号','类别','核查项','结果','详情','建议','参考指导书')
$allRows = @()
foreach ($f in $files) {
    $comp = Get-Component $f.Name
    $rows = Read-SheetData $f.FullName
    foreach ($r in $rows) {
        if ($r.Count -lt 1 -or -not $r[1]) { continue }
        $full = @($comp)
        for ($i = 0; $i -lt 8; $i++) { $full += if ($i -lt $r.Count) { $r[$i] } else { '' } }
        $allRows += ,$full
    }
}
if ($allRows.Count -eq 0) { Write-Host "[跳过] 未提取到数据行"; exit 0 }

# 构建 sheet XML（9 列 A-I）
function Col([int]$n) { [char](64 + $n) }
$sb = New-Object System.Text.StringBuilder
$title = "配置核查汇总报告（共 $($files.Count) 个组件，$($allRows.Count) 项）"
[void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>' + (Xml-Esc $title) + '</t></is></c></row>')
[void]$sb.Append('<row r="4">')
for ($c = 0; $c -lt 9; $c++) {
    $ref = (Col ($c + 1)) + '4'
    [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + $headers[$c] + '</t></is></c>')
}
[void]$sb.Append('</row>')
$ri = 0
foreach ($r in $allRows) {
    $ri++
    $rn = $ri + 4
    [void]$sb.Append('<row r="' + $rn + '">')
    for ($c = 0; $c -lt 9; $c++) {
        $ref = (Col ($c + 1)) + $rn
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + (Xml-Esc $r[$c]) + '</t></is></c>')
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
$outFile = Join-Path (Split-Path -Parent $OutDir) "配置核查汇总报告_$stamp.xlsx"
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

Write-Host "汇总 Excel(.xlsx)报告已生成：$outFile"
