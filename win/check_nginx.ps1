# ============================================================
# 配置核查工具 - Nginx 中间件版（Windows / PowerShell）
# 参考标准：配置核查作业指导书正式版2026_4_1
#           配置核查表_v2.0.0.xlsx（中间件列：Nginx 2项 + 补充Web项6项）
# 运行方式：powershell -ExecutionPolicy Bypass -File check_nginx.ps1
# 输出文件：output\配置核查报告_Nginx_日期时间.html /.xls
# ============================================================
$ErrorActionPreference = 'Continue'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir = Join-Path $ScriptDir "output"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"

# 零依赖 .xlsx 生成器（.NET 自带，无需 Python）
if (Test-Path "$ScriptDir\lib_xlsx.ps1") { . "$ScriptDir\lib_xlsx.ps1" }

$script:R = @()
$script:NGINX_BIN = (Get-Command nginx -ErrorAction SilentlyContinue).Source
$script:NGINX_VERSION = ""

function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += [PSCustomObject]@{ Id=$id; Cat=$cat; Title=$title; Status=$status; Detail=$detail; Chapter=$chapter; Rec=$rec }
}

function Guide-Ref([string]$id) {
    switch -Regex ($id) {
        '^1\.1$'  { return "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" }
        '^1\.14$' { return "《配置核查作业指导书》第1章 系统安全 1.14：中间件应采取限制运行权限和使用安全管理通道等安全加固措施" }
        '^4\.5$'  { return "《配置核查作业指导书》第4章 应用安全 4.5：提供公共信息服务的服务器应具备防DDoS攻击能力" }
        '^4\.6$'  { return "《配置核查作业指导书》第4章 应用安全 4.6：Web应用系统服务器应采取Web防护措施" }
        '^4\.8$'  { return "《配置核查作业指导书》第4章 应用安全 4.8：网站应采取网页防篡改措施，防止对信息内容的非法修改" }
        '^4\.9$'  { return "《配置核查作业指导书》第4章 应用安全 4.9：Web应用系统应具备防范SQL注入、跨站脚本等攻击能力" }
        '^4\.22$' { return "《配置核查作业指导书》第4章 应用安全 4.22：应更改Web应用系统默认服务发布端口" }
        '^4\.23$' { return "《配置核查作业指导书》第4章 应用安全 4.23：应分开设置管理端口与应用端口" }
        default   { return "《配置核查作业指导书》" }
    }
}

function Status-CN([string]$s) {
    switch ($s) { 'pass' { '合规' } 'fail' { '不合规' } 'manual' { '需人工核查' } 'na' { '不适用' } default { $s } }
}
function Status-Color([string]$s) {
    switch ($s) { 'pass' { '#2e7d32' } 'fail' { '#c62828' } 'manual' { '#ef6c00' } 'na' { '#757575' } default { '#000000' } }
}
function Html-Esc([string]$s) {
    $s = $s -replace '&','&amp;'; $s = $s -replace '<','&lt;'; $s = $s -replace '>','&gt;'; return $s
}

# 查找 nginx 配置文件
function Find-NginxConf {
    $candidates = @(
        "$env:USERPROFILE\scoop\apps\nginx\current\conf\nginx.conf",
        "C:\nginx\conf\nginx.conf",
        "C:\Program Files\nginx\conf\nginx.conf"
    )
    foreach ($c in $candidates) { if (Test-Path $c) { return $c } }
    return $null
}
$script:NGINX_CONF = Find-NginxConf

function Check-1_1_Patch {
    $v = $script:NGINX_VERSION
    $mm = ($v -split '/')[-1]
    if (-not $mm) { $mm = $v }
    $maj = ($mm -split '\.')[0]; $min = ($mm -split '\.')[1]
    if ([int]$maj -lt 1 -or ([int]$maj -eq 1 -and [int]$min -lt 24)) {
        Add-Result "1.1" "系统安全" "中间件补丁程序" "fail" "Nginx 当前版本 $v 过旧（<1.24），存在未修复安全漏洞风险。" "第1章" "升级到 Nginx 1.24/1.26 稳定版并安装最新补丁。"
    } else {
        Add-Result "1.1" "系统安全" "中间件补丁程序" "manual" "Nginx 当前版本 $v，请人工核对是否为最新稳定版补丁。" "第1章" "及时升级到最新稳定版本。"
    }
}

function Check-1_14_Hardening {
    $conf = $script:NGINX_CONF
    $issues = @()
    if ($conf) {
        $content = Get-Content $conf -Raw
        # server_tokens
        if ($content -notmatch 'server_tokens\s+off') { $issues += "未隐藏版本号(server_tokens off)" }
        # 目录浏览
        if ($content -match 'autoindex\s+on') { $issues += "开启目录浏览(autoindex on)" }
        # HTTPS（排除注释行）
        $active = ($content -split "`n" | Where-Object { $_ -notmatch '^\s*#' }) -join "`n"
        if ($active -notmatch 'listen\s+443\s+ssl|ssl_certificate\s') { $issues += "未配置HTTPS" }
    } else {
        $issues += "未找到 nginx 配置文件"
    }
    if ($issues.Count -eq 0) {
        Add-Result "1.14" "系统安全" "中间件安全加固" "pass" "Nginx 已做基础安全加固（隐藏版本号、配置 HTTPS、关闭目录浏览）。" "第1章" "持续保持中间件安全配置。"
    } else {
        Add-Result "1.14" "系统安全" "中间件安全加固" "fail" "Nginx 存在以下安全加固缺失：$($issues -join '；')" "第1章" "按缺失项逐条加固中间件配置。"
    }
}

function Check-4_5_Antiddos {
    $conf = $script:NGINX_CONF
    if ($conf -and (Get-Content $conf -Raw) -match 'limit_req|limit_conn') {
        Add-Result "4.5" "应用安全" "防DDoS攻击能力" "pass" "已配置请求/连接限速（limit_req/limit_conn）。" "第4章" "结合前置防火墙/WAF 持续防护 DDoS。"
    } else {
        Add-Result "4.5" "应用安全" "防DDoS攻击能力" "manual" "未检测到 nginx 限速配置，防 DDoS 需结合前置防火墙/WAF，请人工核查。" "第4章" "配置限速或前置 WAF/防火墙防 DDoS。"
    }
}

function Check-4_6_Waf {
    $waf = (& $script:NGINX_BIN -V 2>&1 | Select-String -Pattern 'modsecurity|naxsi|lua' | Select-Object -First 1)
    if ($waf) {
        Add-Result "4.6" "应用安全" "Web防护措施" "pass" "检测到 Web 防护模块：$waf" "第4章" "持续保持 Web 防护规则更新。"
    } else {
        Add-Result "4.6" "应用安全" "Web防护措施" "manual" "未检测到 WAF 模块（modsecurity/naxsi/lua），需结合前置 WAF 防护，请人工核查。" "第4章" "部署 WAF 或前置安全设备。"
    }
}

function Check-4_8_Antitamper {
    $aide = Get-Command aide,tripwire -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($aide) {
        Add-Result "4.8" "应用安全" "网页防篡改" "pass" "检测到网页防篡改/完整性工具：$($aide.Source)" "第4章" "定期校验完整性，防止内容被非法篡改。"
    } else {
        Add-Result "4.8" "应用安全" "网页防篡改" "manual" "未检测到防篡改工具（AIDE/tripwire），请人工核查是否具备网页防篡改能力。" "第4章" "部署 AIDE/tripwire 等完整性校验工具。"
    }
}

function Check-4_9_Sqlinjection {
    Add-Result "4.9" "应用安全" "防范SQL注入/XSS" "manual" "SQL 注入/跨站脚本防护属应用层能力（代码层），需结合 WAF 与代码审计人工核查。" "第4章" "应用层参数化查询 + WAF 双重防护。"
}

function Check-4_22_Defaultport {
    $conf = $script:NGINX_CONF
    $port80 = $null
    if ($conf) {
        $active = (Get-Content $conf | Where-Object { $_ -notmatch '^\s*#' })
        $port80 = $active | Select-String -Pattern 'listen\s+80(\s|;)' | Select-Object -First 1
    }
    if ($port80) {
        Add-Result "4.22" "应用安全" "更改默认服务发布端口" "fail" "仍使用默认 80 端口发布服务：$port80" "第4章" "更改默认 80 端口为非默认端口。"
    } else {
        Add-Result "4.22" "应用安全" "更改默认服务发布端口" "pass" "未使用默认 80 端口发布服务。" "第4章" "持续保持非默认发布端口。"
    }
}

function Check-4_23_Mgmtport {
    Add-Result "4.23" "应用安全" "管理端口与应用端口分离" "na" "Nginx 无独立管理端口（管理通过操作系统），此条不适用于 Nginx。" "第4章" "管理通道与应用通道分离。"
}

# ---------- 报告生成 ----------
function Build-Rows([string]$want) {
    $rows = $script:R | Where-Object { $_.Status -eq $want }
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $rows) {
        $color = Status-Color $r.Status
        $scn = Status-CN $r.Status
        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.Chapter)</td><td>$(Html-Esc $r.Id)</td><td>$(Html-Esc $r.Cat)</td><td>$(Html-Esc $r.Title)</td><td style='color:$color;font-weight:bold;'>$scn</td><td>$(Html-Esc $r.Detail)</td><td>$(Html-Esc $r.Rec)</td><td>$(Html-Esc (Guide-Ref $r.Id))</td></tr>")
    }
    return $sb.ToString()
}

function Gen-Report {
    $pass = ($script:R | Where-Object Status -eq 'pass').Count
    $fail = ($script:R | Where-Object Status -eq 'fail').Count
    $manual = ($script:R | Where-Object Status -eq 'manual').Count
    $na = ($script:R | Where-Object Status -eq 'na').Count
    $meta = "版本：$script:NGINX_VERSION　主机名：$env:COMPUTERNAME　核查时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    $html = @"
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>Nginx 配置核查报告</title>
<style>body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:20px;color:#222;}h1{color:#1a3c6e;}
.meta{background:#f0f4f8;padding:10px 15px;border-radius:6px;margin-bottom:15px;}
.summary{display:flex;gap:15px;margin-bottom:20px;}.card{flex:1;padding:12px;border-radius:6px;text-align:center;color:#fff;}
.card.pass{background:#2e7d32;}.card.fail{background:#c62828;}.card.manual{background:#ef6c00;}.card.na{background:#757575;}
table{border-collapse:collapse;width:100%;font-size:13px;}th,td{border:1px solid #ccc;padding:6px 8px;vertical-align:top;}
th{background:#1a3c6e;color:#fff;position:sticky;top:0;}tr:nth-child(even){background:#f7f9fb;}</style></head><body>
<h1>Nginx 中间件配置核查报告</h1><div class="meta"><div>$meta</div><div>参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx</div></div>
<div class="summary"><div class="card pass">合规<br>$pass</div><div class="card fail">不合规<br>$fail</div><div class="card manual">需人工核查<br>$manual</div><div class="card na">不适用<br>$na</div></div>
"@
    if ($fail -gt 0) { $html += "<h2>一、未通过（$fail 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'fail')</table>" }
    if ($manual -gt 0) { $html += "<h2>二、需人工核查（$manual 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'manual')</table>" }
    if ($na -gt 0) { $html += "<h2>三、不适用（$na 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'na')</table>" }
    if ($pass -gt 0) { $html += "<h2>四、通过（$pass 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'pass')</table>" }
    $html += "</body></html>"

    $htmlfile = Join-Path $OutDir "配置核查报告_Nginx_$Stamp.html"
    $xlsfile = Join-Path $OutDir "配置核查报告_Nginx_$Stamp.xls"
    $html | Out-File -FilePath $htmlfile -Encoding UTF8
    $html | Out-File -FilePath $xlsfile -Encoding UTF8
    Write-Host "HTML报告已生成：$htmlfile"
    Write-Host "Excel(.xls)报告已生成：$xlsfile"
}

# ---------- 主流程 ----------
Write-Host "================================================================"
Write-Host "  配置核查工具 - Nginx 中间件版（Windows/PowerShell）"
Write-Host "  参考标准：配置核查作业指导书正式版2026_4_1"
Write-Host "================================================================"

if (-not $script:NGINX_BIN) {
    Write-Host "  未检测到 nginx，所有项将标记为不适用。"
    Add-Result "1.1" "系统安全" "中间件补丁程序" "na" "未检测到 Nginx，标记不适用。" "第1章" "如部署 Nginx 需及时打补丁。"
    Add-Result "1.14" "系统安全" "中间件安全加固" "na" "未检测到 Nginx，标记不适用。" "第1章" "如部署 Nginx 需做安全加固。"
    Add-Result "4.5" "应用安全" "防DDoS攻击能力" "na" "未检测到 Nginx，标记不适用。" "第4章" "如部署 Nginx 需配置限速。"
    Add-Result "4.6" "应用安全" "Web防护措施" "na" "未检测到 Nginx，标记不适用。" "第4章" "如部署 Nginx 需部署 WAF。"
    Add-Result "4.8" "应用安全" "网页防篡改" "na" "未检测到 Nginx，标记不适用。" "第4章" "如部署 Nginx 需配置防篡改。"
    Add-Result "4.9" "应用安全" "防范SQL注入/XSS" "na" "未检测到 Nginx，标记不适用。" "第4章" "如部署 Nginx 需应用层防护。"
    Add-Result "4.22" "应用安全" "更改默认服务发布端口" "na" "未检测到 Nginx，标记不适用。" "第4章" "如部署 Nginx 需改默认端口。"
    Add-Result "4.23" "应用安全" "管理端口与应用端口分离" "na" "未检测到 Nginx，标记不适用。" "第4章" "管理通道与应用通道分离。"
} else {
    $v = (& $script:NGINX_BIN -v 2>&1 | Out-String)
    if ($v -match 'nginx/([\d.]+)') { $script:NGINX_VERSION = "nginx/$($Matches[1])" }
    Write-Host "  Nginx 版本：$script:NGINX_VERSION"
    Check-1_1_Patch
    Check-1_14_Hardening
    Check-4_5_Antiddos
    Check-4_6_Waf
    Check-4_8_Antitamper
    Check-4_9_Sqlinjection
    Check-4_22_Defaultport
    Check-4_23_Mgmtport
}

$pass = ($script:R | Where-Object Status -eq 'pass').Count
$fail = ($script:R | Where-Object Status -eq 'fail').Count
$manual = ($script:R | Where-Object Status -eq 'manual').Count
$na = ($script:R | Where-Object Status -eq 'na').Count
Write-Host ""
Write-Host "================================================================ "
Write-Host "核查完成，共 $($script:R.Count) 项："
Write-Host "  合规(pass)：$pass    不合规(fail)：$fail    需人工核查(manual)：$manual    不适用(na)：$na"
Gen-Report
if (Get-Command New-Xlsx -ErrorAction SilentlyContinue) {
    New-Xlsx -OutFile "$OutDir\配置核查报告_Nginx_$Stamp.xlsx" -Title "Nginx 中间件配置核查报告" -Rows $script:R -GuideScript { param($id) Guide-Ref $id }
}
Write-Host ""
Write-Host "核查完成，请到 output\ 目录查看 HTML/XLS/XLSX 报告。"
