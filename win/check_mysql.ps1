# ============================================================
# 配置核查工具 - MySQL/MariaDB 版（Windows / PowerShell）
# 参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx（MySQL 19项）
# 运行方式：powershell -ExecutionPolicy Bypass -File check_mysql.ps1
# 连接：默认 mysql -uroot 连 127.0.0.1:3306；可用 $env:MYSQL_PASS 指定口令
# 输出文件：output\配置核查报告_MySQL_日期时间.html /.xls
# ============================================================
$ErrorActionPreference = 'Continue'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir = Join-Path $ScriptDir "output"
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"

# 零依赖 .xlsx 生成器（.NET 自带，无需 Python）
if (Test-Path "$ScriptDir\lib_xlsx.ps1") { . "$ScriptDir\lib_xlsx.ps1" }

$script:R = @()
$script:MYSQL_HOST = '127.0.0.1'
$script:MYSQL_PORT = '3306'
$script:MYSQL_USER = 'root'
$script:MYSQL_PASS = if ($env:MYSQL_PASS) { $env:MYSQL_PASS } else { '' }
$script:DB_VERSION = ""
$script:CONN_OK = $false

function Invoke-Mysql([string]$sql) {
    $args = @('-h', $script:MYSQL_HOST, '-P', $script:MYSQL_PORT, '-u', $script:MYSQL_USER, '-N', '-B')
    if ($script:MYSQL_PASS) { $args += "-p$($script:MYSQL_PASS)" }
    $out = & mysql @args -e $sql 2>$null | Out-String
    return $out.Trim()
}
function Get-MysqlVar([string]$key) {
    $r = Invoke-Mysql "SHOW VARIABLES LIKE '$key';"
    $lines = $r -split "`r?`n"
    if ($lines.Count -ge 1 -and $lines[-1] -match "`t") { return ($lines[-1] -split "`t")[-1].Trim() } else { return $lines[-1].Trim() }
}

function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += [PSCustomObject]@{ Id=$id; Cat=$cat; Title=$title; Status=$status; Detail=$detail; Chapter=$chapter; Rec=$rec }
}
function Guide-Ref([string]$id) {
    switch -Regex ($id) {
        '^1\.1$'  { return "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" }
        '^1\.7$'  { return "《配置核查作业指导书》第1章 系统安全 1.7：数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" }
        '^1\.8$'  { return "《配置核查作业指导书》第1章 系统安全 1.8：数据库管理系统应删除冗余存储过程" }
        '^1\.9$'  { return "《配置核查作业指导书》第1章 系统安全 1.9：数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" }
        '^1\.10$' { return "《配置核查作业指导书》第1章 系统安全 1.10：数据库管理系统应具有自主访问控制功能" }
        '^1\.11$' { return "《配置核查作业指导书》第1章 系统安全 1.11：数据库管理系统应具有备份和恢复功能" }
        '^1\.12$' { return "《配置核查作业指导书》第1章 系统安全 1.12：数据库管理系统应具有表级审计、告警和阻断功能" }
        '^1\.13$' { return "《配置核查作业指导书》第1章 系统安全 1.13：数据库管理系统的数据应和其它应用的数据分类独立存储" }
        '^1\.15$' { return "《配置核查作业指导书》第1章 系统安全 1.15：应具备数据库管理系统超级管理员远程登录限制能力" }
        '^1\.16$' { return "《配置核查作业指导书》第1章 系统安全 1.16：应具备数据库管理系统输入（参数）检查能力" }
        '^1\.19$' { return "《配置核查作业指导书》第1章 系统安全 1.19：应更换数据库管理系统的默认服务端口、管理员用户名和口令" }
        '^1\.20$' { return "《配置核查作业指导书》第1章 系统安全 1.20：数据库管理系统应配置安全策略" }
        '^1\.21$' { return "《配置核查作业指导书》第1章 系统安全 1.21：数据库管理系统应具有行级或列级审计功能" }
        '^1\.22$' { return "《配置核查作业指导书》第1章 系统安全 1.22：数据库管理系统应采取单独、安全监控、审计措施" }
        '^1\.23$' { return "《配置核查作业指导书》第1章 系统安全 1.23：数据库管理系统仅为应用服务器提供访问服务" }
        '^1\.24$' { return "《配置核查作业指导书》第1章 系统安全 1.24：应具备日志审计能力，审计日志至少保留180天" }
        '^1\.25$' { return "《配置核查作业指导书》第1章 系统安全 1.25：检查是否具备边界保护能力，是否可以抗攻击、防篡改" }
        '^1\.26$' { return "《配置核查作业指导书》第1章 系统安全 1.26：检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" }
        '^2\.16$' { return "《配置核查作业指导书》第2章 用户安全 2.16：检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" }
        default   { return "《配置核查作业指导书》" }
    }
}
function Status-CN([string]$s) { switch ($s) { 'pass' { '合规' } 'fail' { '不合规' } 'manual' { '需人工核查' } 'na' { '不适用' } default { $s } } }
function Status-Color([string]$s) { switch ($s) { 'pass' { '#2e7d32' } 'fail' { '#c62828' } 'manual' { '#ef6c00' } 'na' { '#757575' } default { '#000000' } } }
function Html-Esc([string]$s) { $s = $s -replace '&','&amp;'; $s = $s -replace '<','&lt;'; $s = $s -replace '>','&gt;'; return $s }

# 判断版本 EOL
function Mysql-Eol([string]$v) {
    $maj = ($v -split '\.')[0]; $min = ($v -split '\.')[1]
    if ($v -match 'MariaDB') {
        if ([int]$maj -eq 10 -and [int]$min -le 6) { return 'eol' } else { return 'supported' }
    } else {
        if ([int]$maj -le 5) { return 'eol' } else { return 'supported' }
    }
}

# ---------- 核查项 ----------
function Check-1_1 { $e = Mysql-Eol $script:DB_VERSION; if ($e -eq 'eol') { Add-Result "1.1" "系统安全-数据库" "数据库补丁程序" "fail" "版本 $script:DB_VERSION 已 EOL。" "第1章" "升级到受支持版本。" } else { Add-Result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "版本 $script:DB_VERSION，请人工核对最新补丁。" "第1章" "及时升级。" } }
function Check-1_7 {
    $anon = Invoke-Mysql "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='';"
    $empty = Invoke-Mysql "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket');"
    if ($anon) { Add-Result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "存在匿名账户：$($anon -replace "`r?`n",' ')" "第1章" "删除匿名账户。" }
    elseif ($empty) { Add-Result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "存在空口令账户：$($empty -replace "`r?`n",' ')" "第1章" "设置强口令。" }
    else { Add-Result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "未发现匿名/空口令账户。" "第1章" "持续核查账户。" }
}
function Check-1_8 {
    $udf = Invoke-Mysql "SELECT name FROM mysql.func;"
    $risky = ($udf -split "`r?`n") | Where-Object { $_ -match 'sys_exec|sys_eval|sys_bineval|sys_get' }
    if ($risky) { Add-Result "1.8" "系统安全-数据库" "数据库存储过程管理" "fail" "存在高危 UDF：$($risky -join ' ')" "第1章" "删除高危 UDF。" }
    else { Add-Result "1.8" "系统安全-数据库" "数据库存储过程管理" "pass" "未发现高危 UDF。" "第1章" "持续控制。" }
}
function Check-1_9 {
    $danger = Invoke-Mysql "SELECT CONCAT(u.User,'@',u.Host) FROM mysql.user u WHERE u.User NOT IN ('root','mysql.sys','mysql.session','mysql.infoschema') AND (u.File_priv='Y' OR u.Super_priv='Y' OR u.Process_priv='Y' OR u.Shutdown_priv='Y' OR u.Grant_priv='Y' OR (u.Select_priv='Y' AND u.Insert_priv='Y' AND u.Update_priv='Y' AND u.Delete_priv='Y' AND u.Create_priv='Y' AND u.Drop_priv='Y'));"
    if ($danger) { Add-Result "1.9" "系统安全-数据库" "数据库权限最小化" "fail" "存在危险全局权限账户：$($danger -replace "`r?`n",' ')" "第1章" "收敛权限。" }
    else { Add-Result "1.9" "系统安全-数据库" "数据库权限最小化" "pass" "未发现危险全局权限账户。" "第1章" "持续最小权限。" }
}
function Check-1_10 {
    $bind = Get-MysqlVar 'bind_address'; $skip = Get-MysqlVar 'skip_networking'
    if ($skip -eq 'ON') { Add-Result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "skip_networking=ON。" "第1章" "保持。" }
    elseif ($bind -eq '127.0.0.1' -or $bind -eq 'localhost') { Add-Result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "bind_address=$bind。" "第1章" "保持。" }
    else { Add-Result "1.10" "系统安全-数据库" "数据库访问控制" "manual" "bind_address=$bind，请人工确认监听范围。" "第1章" "限制监听地址。" }
}
function Check-1_11 {
    $logbin = Get-MysqlVar 'log_bin'
    if ($logbin -eq 'ON') { Add-Result "1.11" "系统安全-数据库" "数据库备份策略" "manual" "log_bin=ON，但需人工确认备份计划。" "第1章" "建立定期备份。" }
    else { Add-Result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "log_bin=OFF，无备份机制。" "第1章" "启用 binlog + 定期备份。" }
}
function Check-1_12 {
    $glog = Get-MysqlVar 'general_log'
    if ($glog -eq 'ON') { Add-Result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "general_log=ON。" "第1章" "升级为专用审计。" }
    else { Add-Result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "未启用审计（general_log=OFF）。" "第1章" "启用审计。" }
}
function Check-1_13 { Add-Result "1.13" "系统安全-数据库" "数据库与业务隔离" "manual" "数据分类独立存储需人工核查。" "第1章" "分类独立存储。" }
function Check-1_15 {
    $root_remote = Invoke-Mysql "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND LOCATE('%', Host) > 0;"
    $bind = Get-MysqlVar 'bind_address'
    if ($root_remote) { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "fail" "root 允许远程登录：$root_remote" "第1章" "禁止 root 远程。" }
    elseif ($bind -eq '127.0.0.1') { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "已限制（bind=$bind）。" "第1章" "保持。" }
    else { Add-Result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "请人工确认超管远程限制。" "第1章" "限制来源。" }
}
function Check-1_16 { Add-Result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "属应用层能力，需人工核查。" "第1章" "应用层参数化查询。" }
function Check-1_19 {
    $port = Get-MysqlVar 'port'; $root_empty = Invoke-Mysql "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket');"
    $issues = @()
    if ($port -eq '3306') { $issues += "默认端口 3306" }
    if ($root_empty) { $issues += "root 空口令" }
    if ($issues.Count -gt 0) { Add-Result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$($issues -join '；')" "第1章" "改默认端口/口令。" }
    else { Add-Result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/空口令。" "第1章" "持续保持。" }
}
function Check-1_20 {
    $vp = Get-MysqlVar 'validate_password.policy'; $secure = Get-MysqlVar 'secure_file_priv'
    if ($vp -or $secure) { Add-Result "1.20" "系统安全-数据库" "数据库安全策略" "pass" "检测到安全策略（validate_password=$vp，secure_file_priv=$secure）。" "第1章" "完善策略。" }
    else { Add-Result "1.20" "系统安全-数据库" "数据库安全策略" "fail" "未检测到安全策略。" "第1章" "配置安全策略。" }
}
function Check-1_21 { Add-Result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "行/列级审计需人工确认。" "第1章" "配置审计策略。" }
function Check-1_22 { Add-Result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立监控需人工核查。" "第1章" "部署独立监控。" }
function Check-1_23 {
    $app_hosts = Invoke-Mysql "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE Host='%';"
    $bind = Get-MysqlVar 'bind_address'
    if ($app_hosts) { Add-Result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "存在 % 通配账户，请人工确认。" "第1章" "精确限定 Host。" }
    elseif ($bind -eq '127.0.0.1') { Add-Result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "访问受限（bind=$bind）。" "第1章" "保持。" }
    else { Add-Result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "请人工确认。" "第1章" "限制访问。" }
}
function Check-1_24 {
    $expire = Get-MysqlVar 'expire_logs_days'; $sec = Get-MysqlVar 'binlog_expire_logs_seconds'
    $days = 0
    if ($sec -and [int]$sec -gt 0) { $days = [int]($sec / 86400) } elseif ($expire -and [int]$expire -gt 0) { $days = [int]$expire }
    if ($days -ge 180) { Add-Result "1.24" "系统安全-数据库" "日志审计留存180天" "pass" "binlog 保留约 $days 天。" "第1章" "保持。" }
    elseif ($days -gt 0) { Add-Result "1.24" "系统安全-数据库" "日志审计留存180天" "fail" "binlog 保留 $days 天（<180）。" "第1章" "设置 >=180 天。" }
    else { Add-Result "1.24" "系统安全-数据库" "日志审计留存180天" "manual" "请人工核实留存时长。" "第1章" "配置留存。" }
}
function Check-1_25 {
    $fw = Get-NetFirewallRule -Enabled True -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($fw) { Add-Result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到防火墙规则，边界防护需人工核查。" "第1章" "部署边界防护。" }
    else { Add-Result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到防火墙规则。" "第1章" "部署防火墙。" }
}
function Check-1_26 {
    $logerr = Get-MysqlVar 'log_error'; $glog = Get-MysqlVar 'general_log'
    if ($logerr -and $logerr -ne 'stderr') { Add-Result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "pass" "错误日志已配置：$logerr" "第1章" "保持。" }
    else { Add-Result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "fail" "未配置错误日志到独立文件。" "第1章" "配置 log_error。" }
}
function Check-2_16 { $e = Mysql-Eol $script:DB_VERSION; if ($e -eq 'eol') { Add-Result "2.16" "系统安全-数据库" "补丁修复升级到最新" "fail" "版本 $script:DB_VERSION 已 EOL。" "第2章" "升级。" } else { Add-Result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "版本 $script:DB_VERSION，请人工核对最新。" "第2章" "定期升级。" } }

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
    $meta = "连接：$script:MYSQL_USER@$script:MYSQL_HOST`:$script:MYSQL_PORT　版本：$script:DB_VERSION　核查时间：$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $html = @"
<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>MySQL 配置核查报告</title>
<style>body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:20px;color:#222;}h1{color:#1a3c6e;}
.meta{background:#f0f4f8;padding:10px 15px;border-radius:6px;margin-bottom:15px;}
.summary{display:flex;gap:15px;margin-bottom:20px;}.card{flex:1;padding:12px;border-radius:6px;text-align:center;color:#fff;}
.card.pass{background:#2e7d32;}.card.fail{background:#c62828;}.card.manual{background:#ef6c00;}.card.na{background:#757575;}
table{border-collapse:collapse;width:100%;font-size:13px;}th,td{border:1px solid #ccc;padding:6px 8px;vertical-align:top;}
th{background:#1a3c6e;color:#fff;position:sticky;top:0;}tr:nth-child(even){background:#f7f9fb;}</style></head><body>
<h1>MySQL 数据库配置核查报告</h1><div class="meta"><div>$meta</div><div>参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx</div></div>
<div class="summary"><div class="card pass">合规<br>$pass</div><div class="card fail">不合规<br>$fail</div><div class="card manual">需人工核查<br>$manual</div><div class="card na">不适用<br>$na</div></div>
"@
    if ($fail -gt 0) { $html += "<h2>一、未通过（$fail 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'fail')</table>" }
    if ($manual -gt 0) { $html += "<h2>二、需人工核查（$manual 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'manual')</table>" }
    if ($na -gt 0) { $html += "<h2>三、不适用（$na 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'na')</table>" }
    if ($pass -gt 0) { $html += "<h2>四、通过（$pass 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>$(Build-Rows 'pass')</table>" }
    $html += "</body></html>"
    $htmlfile = Join-Path $OutDir "配置核查报告_MySQL_$Stamp.html"
    $xlsfile = Join-Path $OutDir "配置核查报告_MySQL_$Stamp.xls"
    $html | Out-File -FilePath $htmlfile -Encoding UTF8
    $html | Out-File -FilePath $xlsfile -Encoding UTF8
    Write-Host "HTML报告已生成：$htmlfile"
    Write-Host "Excel(.xls)报告已生成：$xlsfile"
}

Write-Host "================================================================"
Write-Host "  配置核查工具 - MySQL/MariaDB 版（Windows/PowerShell）"
Write-Host "  参考标准：配置核查作业指导书正式版2026_4_1"
Write-Host "================================================================"
Write-Host "  连接：$script:MYSQL_USER@$script:MYSQL_HOST`:$script:MYSQL_PORT"

if (-not (Get-Command mysql -ErrorAction SilentlyContinue)) {
    Write-Host "  未检测到 mysql 客户端，所有项将标记为不适用。"
} else {
    $script:DB_VERSION = Invoke-Mysql "SELECT VERSION();"
    if ($script:DB_VERSION) { $script:CONN_OK = $true; Write-Host "  数据库版本：$script:DB_VERSION" }
    else { Write-Host "  [!] 无法连接 MySQL，所有项将标记为需人工核查。" }
}

if ($script:CONN_OK) {
    Check-1_1; Check-1_7; Check-1_8; Check-1_9; Check-1_10; Check-1_11; Check-1_12; Check-1_13; Check-1_15; Check-1_16; Check-1_19; Check-1_20; Check-1_21; Check-1_22; Check-1_23; Check-1_24; Check-1_25; Check-1_26; Check-2_16
} else {
    $titles = @(@("1.1","数据库补丁程序"),@("1.7","数据库账户管理"),@("1.8","数据库存储过程管理"),@("1.9","数据库权限最小化"),@("1.10","数据库访问控制"),@("1.11","数据库备份策略"),@("1.12","数据库审计插件"),@("1.13","数据库与业务隔离"),@("1.15","数据库远程访问控制"),@("1.16","输入验证/SQL注入防护"),@("1.19","数据库默认账户/端口"),@("1.20","数据库安全策略"),@("1.21","数据库操作审计"),@("1.22","数据库审计日志留存"),@("1.23","应用与数据库账户分离"),@("1.24","日志审计留存180天"),@("1.25","边界保护抗攻击防篡改"),@("1.26","防病毒/补丁日志记录"),@("2.16","补丁修复升级到最新"))
    foreach ($it in $titles) {
        $st = if (Get-Command mysql -ErrorAction SilentlyContinue) { 'manual' } else { 'na' }
        Add-Result $it[0] "系统安全-数据库" $it[1] $st "无法连接或未安装 MySQL 客户端。" "第1章" "检查连接参数。"
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
    New-Xlsx -OutFile "$OutDir\配置核查报告_MySQL_$Stamp.xlsx" -Title "MySQL 数据库配置核查报告" -Rows $script:R -GuideScript { param($id) Guide-Ref $id }
}
Write-Host ""
Write-Host "核查完成，请到 output\ 目录查看 HTML/XLS/XLSX 报告。"
