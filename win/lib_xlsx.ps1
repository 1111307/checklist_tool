# ============================================================
# 零依赖真 .xlsx 生成器（PowerShell + .NET，无需 Python/openpyxl）
# ============================================================
# 原理：.xlsx 本质是 ZIP 包内若干 XML 文件，用 .NET System.IO.Compression 即可打包。
# 用法（在核查脚本内 dot-source 后调用）：
#   . "$ScriptDir\lib_xlsx.ps1"
#   New-Xlsx -OutFile "$OutDir\配置核查报告_MySQL_$Stamp.xlsx" -Title "MySQL 数据库配置核查报告" -Rows $script:R -GuideScript { param($id) Guide-Ref $id }
# ============================================================

function Xml-Esc([string]$s) {
    if ($null -eq $s) { return '' }
    $s = $s -replace '&', '&amp;'
    $s = $s -replace '<', '&lt;'
    $s = $s -replace '>', '&gt;'
    $s = $s -replace '"', '&quot;'
    $s = $s -replace "'", '&apos;'
    # 换行转为数字引用，Excel 内可换行显示
    $s = $s -replace "`r`n", '&#10;' -replace "`n", '&#10;' -replace "`r", ''
    return $s
}

function New-Xlsx {
    param(
        [string]$OutFile,
        [string]$Title = '配置核查报告',
        [object[]]$Rows,
        [scriptblock]$GuideScript = $null
    )

    # 统计
    $pass = 0; $fail = 0; $manual = 0; $na = 0
    foreach ($r in $Rows) {
        switch ($r.Status) { 'pass' { $pass++ } 'fail' { $fail++ } 'manual' { $manual++ } 'na' { $na++ } }
    }
    function Local-StatusCN([string]$s) { switch ($s) { 'pass' { '合规' } 'fail' { '不合规' } 'manual' { '需人工核查' } 'na' { '不适用' } default { $s } } }

    # 构建 sheetData
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<row r="1"><c r="A1" t="inlineStr"><is><t>' + (Xml-Esc $Title) + '</t></is></c></row>')
    $summary = "合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na　共 $($Rows.Count) 项"
    [void]$sb.Append('<row r="2"><c r="A2" t="inlineStr"><is><t>' + (Xml-Esc $summary) + '</t></is></c></row>')
    $hd = @('章节','编号','类别','核查项','结果','详情','建议','参考指导书')
    [void]$sb.Append('<row r="4">')
    for ($c = 0; $c -lt 8; $c++) {
        $ref = [char](65 + $c) + '4'
        [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + $hd[$c] + '</t></is></c>')
    }
    [void]$sb.Append('</row>')
    $ri = 0
    foreach ($r in $Rows) {
        $ri++
        $rn = $ri + 4
        $guide = if ($GuideScript) { & $GuideScript $r.Id } else { '' }
        $vals = @($r.Chapter, $r.Id, $r.Cat, $r.Title, (Local-StatusCN $r.Status), $r.Detail, $r.Rec, $guide)
        [void]$sb.Append('<row r="' + $rn + '">')
        for ($c = 0; $c -lt 8; $c++) {
            $ref = [char](65 + $c) + $rn
            [void]$sb.Append('<c r="' + $ref + '" t="inlineStr"><is><t>' + (Xml-Esc $vals[$c]) + '</t></is></c>')
        }
        [void]$sb.Append('</row>')
    }
    $sheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + $sb.ToString() + '</sheetData></worksheet>'

    # 其余 XML
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' + "`n" + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' + "`n" + '<Default Extension="xml" ContentType="application/xml"/>' + "`n" + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' + "`n" + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' + "`n" + '</Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' + "`n" + '</Relationships>'
    $wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + "`n" + '<sheets><sheet name="核查明细" sheetId="1" r:id="rId1"/></sheets>' + "`n" + '</workbook>'
    $wbrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + "`n" + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + "`n" + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' + "`n" + '</Relationships>'

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $enc = New-Object System.Text.UTF8Encoding($false)
    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }

    $fs = [System.IO.File]::Create($OutFile)
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

    Write-Host "Excel(.xlsx)报告已生成：$OutFile"
}
