# ============================================================
# 配置核查工具 - Redis 版（Windows / PowerShell）
# 参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx（Redis 16项）
# 运行方式：powershell -ExecutionPolicy Bypass -File check_redis.ps1
# 连接：默认 redis-cli 连 127.0.0.1:6379；可用 $env:REDIS_PASS 指定口令
# 输出文件：output\配置核查报告_Redis_日期时间.html /.xls
# ============================================================
$ErrorActionPreference = 'Continue'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir = Join-Path $ScriptDir "output"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"

# 零依赖 .xlsx 生成器（.NET 自带，无需 Python）
if (Test-Path "$ScriptDir\lib_xlsx.ps1") { . "$ScriptDir\lib_xlsx.ps1" }

$script:R = @()
$script:REDIS_HOST = '127.0.0.1'
$script:REDIS_PORT = '6379'
$script:REDIS_PASS = if ($env:REDIS_PASS) { $env:REDIS_PASS } else { '' }
$script:REDIS_VERSION = ""
$script:CONN_OK = $false

function Invoke-Redis([string]$cmd) {
    $args = @('-h', $script:REDIS_HOST, '-p', $script:REDIS_PORT, '--no-auth-warning')
    if ($script:REDIS_PASS) { $args += @('-a', $script:REDIS_PASS) }
    $out = & redis-cli @args $cmd.Split(' ') 2>$null | Out-String
    return $out.Trim()
}
function Get-RedisCfg([string]$key) {
    $r = Invoke-Redis "CONFIG GET $key"
    $lines = $r -split "`r?`n"
    if ($lines.Count -ge 2) { return $lines[-1].Trim() } else { return '' }
}

function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += [PSCustomObject]@{ Id=$id; Cat=$cat; Title=$title; Status=$status; Detail=$detail; Chapter=$chapter; Rec=$rec }
}
function Guide-Ref([string]$id) {
    switch -Regex ($id) {
        '^1\.1$'  { return "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" }
        '^1\.7$'  { return "《配置核查作业指导书》第1章 系统安全 1.7：数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" }
        '^1\.8$'  { return "《配置核查作业指导书》第1章 系统安全 1.8：数据库管理系统应删除冗余存储过程" }
        '^1\.10$' { return "《配置核查作业指导书》第1章 系统安全 1.10：数据库管理系统应具有自主访问控制功能" }
        '^1\.11$' { return "《配置核查作业指导书》第1章 系统安全 1.11：数据库管理系统应具有备份和恢复功能" }
        '^1\.12$' { return "《配置核查作业指导书》第1章 系统安全 1.12：数据库管理系统应具有表级审计、告警和阻断功能" }
        '^1\.15$' { return "《配置核查作业指导书》第1章 系统安全 1.15：应具备数据库管理系统超级管理员远程登录限制能力" }
        '^1\.16$' { return "《配置核查作业指导书》第1章 系统安全 1.16：应具备数据库管理系统输入（参数）检查能力" }
        '^1\.19$' { return "《配置核查作业指导书》第1章 系统安全 1.19：应更换数据库管理系统的默认服务端口、管理员用户名和口令" }
        '^1\.20$' { return "《配置核查作业指导书》第1章 系统安全 1.20：数据库管理系统应配置安全策略" }
        '^1\.21$' { return "《配置核查作业指导书》第1章 系统安全 1.21：数据库管理系统应具有行级或列级审计功能" }
        '^1\.22$' { return "《配置核查作业指导书》第1章 系统安全 1.22：数据库管理系统应采取单独、安全监控、审计措施" }
        '^1\.23$' { return "《配置核查作业指导书》第1章 系统安全 1.23：数据库管理系统仅为应用服务器提供访问服务" }
        '^1\.25$' { return "《配置核查作业指导书》第1章 系统安全 1.25：检查是否具备边界保护能力，是否可以抗攻击、防篡改" }
        '^1\.26$' { return "《配置核查作业指导书》第1章 系统安全 1.26：检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" }
        '^2\.16$' { return "《配置核查作业指导书》第2章 用户安全 2.16：检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" }
        default   { return "《配置核查作业指导书》" }
    }
}
function Status-CN([string]$s) { switch ($s) { 'pass' { '合规' } 'fail' { '不合规' } 'manual' { '需人工核查' } 'na' { '不适用' } default { $s } } }
function Status-Color([string]$s) { switch ($s) { 'pass' { '#2e7d32' } 'fail' { '#c62828' } 'manual' { '#ef6c00' } 'na' { '#757575' } default { '#000000' } } }
function Html-Esc([string]$s) { $s = $s -replace '&','&amp;'; $s = $s -replace '<','&lt;'; $s = $s -replace '>','&gt;'; return $s }

# 从 INFO server 取 redis_version
function Get-RedisVersion {
    $info = Invoke-Redis "INFO server"
    if ($info -match 'redis_version:([\d.]+)') { return $Matches[1] }
    return ''
}

# ---------- 核查项 ----------
function Check-1_1_Patch {
    $v = $script:REDIS_VERSION
    $maj = ($v -split '\.')[0]
    if ($v -and [int]$maj -lt 7) {
        Add-Result "1.1" "系统安全-数据库" "数据库补丁程序" "fail" "Redis 版本 $v 已停止维护（<7.0）。" "第1章" "升级到 Redis 7.x/8.x。"
    } else {
        Add-Result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "Redis 版本 $v（仍在维护期），请人工核对最新补丁。" "第1章" "及时升级到最新稳定版本。"
    }
}
function Check-1_7_Auth {
    $pass = Get-RedisCfg 'requirepass'
    if (-not $pass) { Add-Result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "未设置 requirepass，Redis 无需口令即可访问。" "第1章" "设置 requirepass 强口令。" }
    else { Add-Result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "已设置 requirepass 认证口令。" "第1章" "持续使用强口令并定期更换。" }
}
function Check-1_8_Dangerous {
    $risky = @()
    foreach ($c in 'FLUSHALL','FLUSHDB','EVAL','SHUTDOWN') {
        $info = Invoke-Redis "COMMAND INFO $c"
        if ($info -and $info -notmatch '^\s*$') { $risky += $c }
    }
    if ($risky.Count -gt 0) { Add-Result "1.8" "系统安全-数据库" "数据库存储过程管理" "fail" "以下高危命令未禁用：$($risky -join ' ')" "第1章" "rename-command 禁用高危命令。" }
    else { Add-Result "1.8" "系统安全-数据库" "数据库存储过程管理" "pass" "高危命令已重命名或禁用。" "第1章" "持续禁用不必要的危险命令。" }
}
function Check-1_10_Access {
    $pm = Get-RedisCfg 'protected-mode'; $bind = Get-RedisCfg 'bind'
    if ($pm -eq 'no' -and (-not $bind -or $bind -eq '0.0.0.0')) { Add-Result "1.10" "系统安全-数据库" "数据库访问控制" "fail" "protected-mode=no 且 bind 为空/0.0.0.0，对全网开放。" "第1章" "开启 protected-mode 并绑定地址。" }
    else { Add-Result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "protected-mode=$pm，bind=$bind，访问受限。" "第1章" "持续限制监听地址。" }
}
function Check-1_11_Backup {
    $save = Get-RedisCfg 'save'; $aof = Get-RedisCfg 'appendonly'
    if ($save -and $save -ne '') { Add-Result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "已配置 RDB 快照（save=$save）。" "第1章" "结合 RDB/AOF 定期备份。" }
    elseif ($aof -eq 'yes') { Add-Result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "appendonly=yes。" "第1章" "结合 RDB/AOF 定期备份。" }
    else { Add-Result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "未配置持久化（save 空且 appendonly=no）。" "第1章" "开启 RDB/AOF 持久化。" }
}
function Check-1_12_Audit {
    $slow = Get-RedisCfg 'slowlog-log-slower-than'
    if ($slow -and $slow -ne '-1') { Add-Result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "slowlog 已启用。" "第1章" "结合慢日志与外部审计。" }
    else { Add-Result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "slowlog 未启用。" "第1章" "开启 slowlog。" }
}
function Check-1_15_Remote {
    $pass = Get-RedisCfg 'requirepass'; $bind = Get-RedisCfg 'bind'; $pm = Get-RedisCfg 'protected-mode'
    if (-not $pass) { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "fail" "未设置口令，任意来源可远程执行管理命令。" "第1章" "设置口令并限制来源。" }
    elseif ($bind -eq '127.0.0.1' -or $pm -eq 'yes') { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "已设口令且访问受限。" "第1章" "持续限制远程来源。" }
    else { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "已设口令但 bind=$bind，请人工确认限制。" "第1章" "通过防火墙/绑定限制来源。" }
}
function Check-1_16_Input { Add-Result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "Redis 输入检查属应用层能力，需人工核查。" "第1章" "应用层校验输入。" }
function Check-1_19_Defaults {
    $port = Get-RedisCfg 'port'; $pass = Get-RedisCfg 'requirepass'
    $issues = @()
    if ($port -eq '6379') { $issues += "默认端口 6379" }
    if (-not $pass) { $issues += "未设置口令" }
    if ($issues.Count -gt 0) { Add-Result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$($issues -join '；')" "第1章" "改默认端口、设强口令。" }
    else { Add-Result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/无口令问题。" "第1章" "持续保持。" }
}
function Check-1_20_Policy {
    $maxmem = Get-RedisCfg 'maxmemory'; $pm = Get-RedisCfg 'protected-mode'
    if (($maxmem -and $maxmem -ne '0') -or $pm -eq 'yes') { Add-Result "1.20" "系统安全-数据库" "数据库安全策略" "pass" "检测到安全策略（maxmemory=$maxmem，protected-mode=$pm）。" "第1章" "完善安全基线。" }
    else { Add-Result "1.20" "系统安全-数据库" "数据库安全策略" "fail" "未检测到明显安全策略。" "第1章" "配置 maxmemory 等策略。" }
}
function Check-1_21_Audit { Add-Result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "Redis 无行级/列级审计，需结合外部审计工具。" "第1章" "结合外部审计。" }
function Check-1_22_Monitor { Add-Result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立监控审计需人工核查。" "第1章" "部署独立监控。" }
function Check-1_23_Appsep {
    $bind = Get-RedisCfg 'bind'; $pm = Get-RedisCfg 'protected-mode'
    if ($bind -eq '127.0.0.1' -or $pm -eq 'yes') { Add-Result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "访问受限（bind=$bind）。" "第1章" "持续确保仅应用服务器访问。" }
    else { Add-Result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "请人工确认是否仅应用服务器访问。" "第1章" "绑定内网地址。" }
}
function Check-1_25_Boundary {
    $fw = Get-NetFirewallRule -Enabled True -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($fw) { Add-Result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到 Windows 防火墙规则，但边界防护需人工核查。" "第1章" "部署边界防护。" }
    else { Add-Result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到防火墙规则。" "第1章" "部署防火墙。" }
}
function Check-1_26_Log {
    $logfile = Get-RedisCfg 'logfile'
    if ($logfile -and $logfile -ne '""' -and $logfile -ne '') { Add-Result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "pass" "已配置日志文件：$logfile" "第1章" "确保日志完整。" }
    else { Add-Result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "fail" "logfile 未配置到独立文件。" "第1章" "配置 logfile。" }
}
function Check-2_16_Patch {
    $v = $script:REDIS_VERSION; $maj = ($v -split '\.')[0]
    if ($v -and [int]$maj -lt 7) { Add-Result "2.16" "系统安全-数据库" "补丁修复升级到最新" "fail" "Redis 版本 $v 已停止维护。" "第2章" "升级到 Redis 7.x/8.x。" }
    else { Add-Result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "Redis 版本 $v，请人工核对最新。" "第2章" "定期升级。" }
}

function Build-Rows([string]$want) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in ($script:R | Where-Object { $_.Status -eq $want })) {
        $color = Status-Color $r.Status; $scn = Status-CN $r.Status
        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.Chapter)</td><td>$(Html-Esc $r.Id)</td><td>$(Html-Esc $r.Cat)</td><td>$(Html-Esc $r.Title)</td><td style='color:$color;font-weight:bold;'>$scn</td><td>$(Html-Esc $r.Detail)</td><td>$(Html-Esc $r.Rec)</td><td>$(Html-Esc (Guide-Ref $r.Id))</td></tr>")
    }
    return $sb.ToString()
}
function Gen-Report {
    $pass = ($script:R | Where-Object Status -eq 'pass').Count
    $fail = ($script:R | Where-Object Status -eq 'fail').Count
    $manual = ($script:R | Where-Object Status -eq 'manual').Count
    $na = ($script:R | Where-Object Status -eq 'na').Count
    $meta = "连接：$script:REDIS_HOST`:$script:REDIS_PORT　版本：$script:REDIS_VERSION　主机名：$env:COMPUTERNAME　核查时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $html = @"
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>Redis 配置核查报告</title>
<style>body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:20px;color:#222;}h1{color:#1a3c6e;}
.meta{background:#f0f4f8;padding:10px 15px;border-radius:6px;margin-bottom:15px;}
.summary{display:flex;gap:15px;margin-bottom:20px;}.card{flex:1;padding:12px;border-radius:6px;text-align:center;color:#fff;}
.card.pass{background:#2e7d32;}.card.fail{background:#c62828;}.card.manual{background:#ef6c00;}.card.na{background:#757575;}
table{border-collapse:collapse;width:100%;font-size:13px;}th,td{border:1px solid #ccc;padding:6px 8px;vertical-align:top;}
th{background:#1a3c6e;color:#fff;position:sticky;top:0;}tr:nth-child(even){background:#f7f9fb;}</style></head><body>
<h1>Redis 配置核查报告</h1><div class="meta"><div>$meta</div><div>参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx</div></div>
<div class="summary"><div class="card pass">合规<br>$pass</div><div class="card fail">不合规<br>$fail</div><div class="card manual">需人工核查<br>$manual</div><div class="card na">不适用<br>$na</div></div>
"@
    if ($fail -gt 0) { $html += "<h2>一、未通过（$fail 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'fail')</table>" }
    if ($manual -gt 0) { $html += "<h2>二、需人工核查（$manual 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'manual')</table>" }
    if ($na -gt 0) { $html += "<h2>三、不适用（$na 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'na')</table>" }
    if ($pass -gt 0) { $html += "<h2>四、通过（$pass 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'pass')</table>" }
    $html += "</body></html>"
    $htmlfile = Join-Path $OutDir "配置核查报告_Redis_$Stamp.html"
    $xlsfile = Join-Path $OutDir "配置核查报告_Redis_$Stamp.xls"
    $html | Out-File -FilePath $htmlfile -Encoding UTF8
    $html | Out-File -FilePath $xlsfile -Encoding UTF8
    Write-Host "HTML报告已生成：$htmlfile"
    Write-Host "Excel(.xls)报告已生成：$xlsfile"
}

Write-Host "================================================================"
Write-Host "  配置核查工具 - Redis 版（Windows/PowerShell）"
Write-Host "  参考标准：配置核查作业指导书正式版2026_4_1"
Write-Host "================================================================"
Write-Host "  连接：$script:REDIS_HOST`:$script:REDIS_PORT"

if (-not (Get-Command redis-cli -ErrorAction SilentlyContinue)) {
    Write-Host "  未检测到 redis-cli，所有项将标记为不适用。"
    foreach ($it in @(@("1.1","数据库补丁程序"),@("1.7","数据库账户管理"),@("1.8","数据库存储过程管理"),@("1.10","数据库访问控制"),@("1.11","数据库备份策略"),@("1.12","数据库审计插件"),@("1.15","数据库远程访问控制"),@("1.16","输入验证/SQL注入防护"),@("1.19","数据库默认账户/端口"),@("1.20","数据库安全策略"),@("1.21","数据库操作审计"),@("1.22","数据库审计日志留存"),@("1.23","应用与数据库账户分离"),@("1.25","边界保护抗攻击防篡改"),@("1.26","防病毒/补丁日志记录"),@("2.16","补丁修复升级到最新"))) {
        Add-Result $it[0] "系统安全-数据库" $it[1] "na" "未检测到 redis-cli，标记不适用。" "第1章" "安装 redis-cli。"
    }
} else {
    $script:REDIS_VERSION = Get-RedisVersion
    if ($script:REDIS_VERSION) { $script:CONN_OK = $true; Write-Host "  Redis 版本：$script:REDIS_VERSION" }
    else { Write-Host "  [!] 无法连接 Redis（$script:REDIS_HOST`:$script:REDIS_PORT），所有项将标记为需人工核查。" }
    if ($script:CONN_OK) {
        Check-1_1_Patch; Check-1_7_Auth; Check-1_8_Dangerous; Check-1_10_Access; Check-1_11_Backup; Check-1_12_Audit; Check-1_15_Remote; Check-1_16_Input; Check-1_19_Defaults; Check-1_20_Policy; Check-1_21_Audit; Check-1_22_Monitor; Check-1_23_Appsep; Check-1_25_Boundary; Check-1_26_Log; Check-2_16_Patch
    } else {
        foreach ($it in @(@("1.1","数据库补丁程序"),@("1.7","数据库账户管理"),@("1.8","数据库存储过程管理"),@("1.10","数据库访问控制"),@("1.11","数据库备份策略"),@("1.12","数据库审计插件"),@("1.15","数据库远程访问控制"),@("1.16","输入验证/SQL注入防护"),@("1.19","数据库默认账户/端口"),@("1.20","数据库安全策略"),@("1.21","数据库操作审计"),@("1.22","数据库审计日志留存"),@("1.23","应用与数据库账户分离"),@("1.25","边界保护抗攻击防篡改"),@("1.26","防病毒/补丁日志记录"),@("2.16","补丁修复升级到最新"))) {
            Add-Result $it[0] "系统安全-数据库" $it[1] "manual" "无法连接 Redis，需人工核查。" "第1章" "检查连接参数。"
        }
    }
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
    New-Xlsx -OutFile "$OutDir\配置核查报告_Redis_$Stamp.xlsx" -Title "Redis 配置核查报告" -Rows $script:R -GuideScript { param($id) Guide-Ref $id }
}
Write-Host ""
Write-Host "核查完成，请到 output\ 目录查看 HTML/XLS/XLSX 报告。"
