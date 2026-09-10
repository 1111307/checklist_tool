# ============================================================
# 配置核查工具 - 网络设备版（Windows / PowerShell，华为/华三/锐捷）
# 参考标准：配置核查作业指导书v2.2（第5章 网络安全 5.1-5.23）
#
# 与麒麟版 kylin/check_network.sh 同判定逻辑（Win7+ / PowerShell 2.0+）。
# 网络设备为独立硬件，本脚本采用「采集-解析」两步模式：
#   1) powershell -ExecutionPolicy Bypass -File check_network.ps1 init
#      生成采集工作区与三厂商命令清单模板——运维陪同登录各设备，
#      按模板粘贴命令回显，每台设备存一个 txt
#   2) powershell -ExecutionPolicy Bypass -File check_network.ps1 check
#      解析回显文件，按 23 项判定，生成报告
#   （或直接双击 check_network.bat init / check）
#
# 采集文件约定：output\netdev\<设备名>.txt，各命令段以
#   ########## CMD: <命令> ########## 行分隔（模板已带标记）
# 输出：output\配置核查报告_网络设备_日期时间.html / .xls
# ============================================================
param([string]$Command = "")
$ErrorActionPreference = 'Continue'

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OutDir = Join-Path $ScriptDir "output"
$NetDir = Join-Path $OutDir "netdev"
$Stamp = Get-Date -Format "yyyyMMdd_HHmmss"

# ---------- 结果存储 ----------
$script:R = @()
function Add-Result([string]$id, [string]$cat, [string]$title, [string]$status, [string]$detail, [string]$chapter, [string]$rec) {
    $script:R += @{ id = $id; cat = $cat; title = $title; status = $status; detail = $detail; chapter = $chapter; rec = $rec; guide = "《配置核查作业指导书v2.2》第5章 网络安全 $id" }
}
function Html-Esc([string]$s) {
    if ($null -eq $s) { return "" }
    $s -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;'
}

# ---------- 文本工具 ----------
# 设备文件编码自适应：先按 UTF-8 严格解码，失败按 GBK（记事本默认 ANSI）
function Read-Text([string]$path) {
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
    try { return $strictUtf8.GetString($bytes) }
    catch { return [System.Text.Encoding]::GetEncoding(936).GetString($bytes) }
}
# bash grep -c 等价：统计匹配模式的行数（$ignoreCase 对应 grep -i）
function Count-Lines([string]$text, [string]$pattern, [bool]$ignoreCase) {
    $c = 0
    if ([string]::IsNullOrEmpty($text)) { return 0 }
    foreach ($line in ($text -split "`r?`n")) {
        if ($ignoreCase) { if ($line -imatch $pattern) { $c++ } }
        else { if ($line -cmatch $pattern) { $c++ } }
    }
    return $c
}
function NonEmpty([string]$text) {
    return ((($text -replace '\s', '') -ne ''))
}
# 统计 VLAN 条数：优先数 VLAN 表格行（华为/华三/锐捷 brief 输出），
# 退化用汇总行 "The total number of vlans is : N"，再退化旧写法 "VLAN N"
function Count-Vlans([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return 0 }
    $n = Count-Lines $text '^[0-9]+\s+\S' $false
    if ($n -eq 0) {
        $m = [regex]::Match($text, '(?i)total number of vlans?[^0-9]*([0-9]+)')
        if ($m.Success) { $n = [int]$m.Groups[1].Value }
    }
    if ($n -eq 0) { $n = Count-Lines $text '^(VLAN|vlan) [0-9]+' $false }
    return $n
}
# 筛出端口状态表里的接口行：兼容全称 GigabitEthernet0/0/1、缩写 GE0/0/1、
# 华三 XGE1/0/1、锐捷 connected/notconnect 四种写法
function Get-IfLines([string]$text) {
    if ([string]::IsNullOrEmpty($text)) { return @() }
    return @($text -split "`r?`n" | Where-Object {
        $_ -imatch '^[A-Za-z][A-Za-z0-9./:-]*[0-9](/[0-9]+)*\s+(up|down|connected|notconnect|dormant|inactive|active|suspended)'
    })
}

# ---------- init：生成采集模板 ----------
function Do-Init {
    New-Item -ItemType Directory -Force -Path $NetDir | Out-Null
    $utf8 = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $NetDir "采集说明.txt"), @'
【网络设备配置核查 - 采集操作说明】
1. 每台网络设备（交换机/路由器/防火墙）建一个 txt 文件，文件名即设备名，
   如：核心交换机-01.txt、边界防火墙-01.txt（保存为 UTF-8 或 ANSI 均可）。
2. 复制 采集命令清单-华为华三.txt 或 采集命令清单-锐捷.txt 的全部内容到该文件，
   登录设备后逐条执行命令，把每条命令的完整回显粘贴到对应 ########## CMD: 行之后。
   （在设备上可用 terminal monitor / screen-length 0 temporary 之类命令取消分页）
3. 全部设备采集完成后，运行：check_network.bat check 生成核查报告。
4. 只采集了部分设备也能出报告，未覆盖项将标记需人工核查。
'@, $utf8)
    [System.IO.File]::WriteAllText((Join-Path $NetDir "采集命令清单-华为华三.txt"), @'
########## CMD: display version ##########

########## CMD: display current-configuration ##########

########## CMD: display vlan ##########

########## CMD: display ip routing-table ##########

########## CMD: display ssh server status ##########

########## CMD: display local-user ##########

########## CMD: display acl all ##########

########## CMD: display dot1x ##########

########## CMD: display port-isolate group ##########

########## CMD: display multicast routing-table ##########

########## CMD: display interface brief ##########

########## CMD: display users ##########

（防火墙设备补充执行以下命令）
########## CMD: display firewall session table ##########

########## CMD: display ike sa ##########

########## CMD: display hrp state ##########

'@, $utf8)
    [System.IO.File]::WriteAllText((Join-Path $NetDir "采集命令清单-锐捷.txt"), @'
########## CMD: show version ##########

########## CMD: show running-config ##########

########## CMD: show vlan ##########

########## CMD: show ip route ##########

########## CMD: show ip ssh ##########

########## CMD: show users ##########

########## CMD: show access-lists ##########

########## CMD: show dot1x summary ##########

########## CMD: show interfaces status ##########

（防火墙设备补充执行以下命令）
########## CMD: show session ##########

'@, $utf8)
    Write-Host "采集工作区已生成：$NetDir\"
    Write-Host "  采集说明.txt                    —— 先读这个"
    Write-Host "  采集命令清单-华为华三.txt        —— 华为/华三设备命令模板"
    Write-Host "  采集命令清单-锐捷.txt            —— 锐捷设备命令模板"
    Write-Host "按说明采集完成后运行：check_network.bat check"
}

# ---------- check：解析回显并判定 ----------
$script:AllText = ""
$script:DevCount = 0
$script:DevNames = ""

function Get-Sec([string]$cmdKeyword) {
    # 提取全部设备回显中 "CMD: <关键字>" 段（命令标记行之后到下一个 CMD 标记行）
    $marker = "CMD: $cmdKeyword"
    $sb = New-Object System.Text.StringBuilder
    $on = $false
    foreach ($line in ($script:AllText -split "`r?`n")) {
        if ($line.IndexOf($marker) -ge 0) { $on = $true; continue }
        if ($line -match '^\s*#+\s*CMD:') { $on = $false; continue }
        if ($on) { [void]$sb.AppendLine($line) }
    }
    return $sb.ToString()
}

function Load-Devices {
    if (-not (Test-Path $NetDir)) {
        Write-Host "未找到采集目录 $NetDir，请先运行：check_network.bat init"
        exit 1
    }
    $names = @()
    $files = Get-ChildItem -Path (Join-Path $NetDir "*.txt") -ErrorAction SilentlyContinue
    foreach ($f in $files) {
        if ($f.Name -like "采集说明*" -or $f.Name -like "采集命令清单*") { continue }
        $content = Read-Text $f.FullName
        if ($content -match 'CMD:') {
            $script:DevCount++
            $names += $f.BaseName
            $script:AllText += "`n" + $content
        }
    }
    if ($script:DevCount -eq 0) {
        Write-Host "采集目录中没有含 CMD: 标记的设备回显文件（当前目录：$NetDir）。"
        Write-Host "请按 采集说明.txt 将命令回显粘贴进模板文件后再运行 check。"
        exit 1
    }
    $script:DevNames = ($names -join " ")
    Write-Host "已加载 $($script:DevCount) 台设备回显：$($script:DevNames)"
}

function Do-Check {
    Load-Devices

    $cfg    = (Get-Sec 'display current-configuration') + (Get-Sec 'show running-config')
    $vlan   = (Get-Sec 'display vlan') + (Get-Sec 'show vlan')
    $route  = (Get-Sec 'display ip routing-table') + (Get-Sec 'show ip route')
    $sshSt  = (Get-Sec 'display ssh server status') + (Get-Sec 'show ip ssh')
    $acl    = (Get-Sec 'display acl all') + (Get-Sec 'show access-lists')
    $dot1x  = (Get-Sec 'display dot1x') + (Get-Sec 'show dot1x summary')
    $pisol  = Get-Sec 'display port-isolate group'
    $mcast  = Get-Sec 'display multicast routing-table'
    $brief  = (Get-Sec 'display interface brief') + (Get-Sec 'show interfaces status')
    $ike    = Get-Sec 'display ike sa'
    $hrp    = Get-Sec 'display hrp state'
    $ver    = (Get-Sec 'display version') + (Get-Sec 'show version')

    # ---- 5.1 跨网跨域交换流程（文档类） ----
    Add-Result "5.1" "网络安全" "跨网跨域数据交换按规定流程进行" "manual" `
        "文档流程类核查：调阅跨网跨域交换审批单与方案评审意见，比对设备侧交换通道与审批范围（见指导书 5.1 方法（4）日志抽查）。" `
        "第5章" "跨网跨域交换须逐次审批并留存方案评审记录。"
    # ---- 5.2 无线组网流程（文档类） ----
    Add-Result "5.2" "网络安全" "无线网络构建按规定流程执行" "manual" `
        "文档流程类核查：查阅无线建设审批、方案评审，并在授权下探测是否存在未登记私建热点（见指导书 5.2 方法（4））。" `
        "第5章" "无线组网须审批后实施，私建热点应立即清除。"
    # ---- 5.3 最小化架构 ----
    if (NonEmpty $vlan) {
        $vlanCnt = Count-Vlans $vlan
        if ($vlanCnt -le 1) {
            Add-Result "5.3" "网络安全" "按最小化原则设计网络架构" "fail" `
                "设备回显中仅发现 $vlanCnt 个 VLAN，疑似未按业务划分区域（全设备置于同一广播域）。请核对 topology 与规划表。" `
                "第5章" "按业务划分 VLAN，采用交换技术组网并最小化广播域。"
        } else {
            Add-Result "5.3" "网络安全" "按最小化原则设计网络架构" "pass" `
                "回显发现 $vlanCnt 个 VLAN，已按业务划分。请结合 IP 地址规划表人工确认无整段闲置。" `
                "第5章" "定期复核 VLAN 与 IP 规划的最小化。"
        }
    } else {
        Add-Result "5.3" "网络安全" "按最小化原则设计网络架构" "manual" `
            "未采集到 display vlan / show vlan 回显，无法自动判定。" "第5章" "补采 VLAN 配置后重跑，或按指导书 5.3 人工核查。"
    }
    # ---- 5.4 边界互联与路由最小化 ----
    if (NonEmpty $route) {
        $protos = @([regex]::Matches($route, 'OSPF|RIP|BGP|ISIS') | ForEach-Object { $_.Value } | Select-Object -Unique)
        $protoCnt = $protos.Count
        if ($protoCnt -gt 2) {
            Add-Result "5.4" "网络安全" "边界互联节点与路由最小化" "fail" `
                "路由表中同时出现 $protoCnt 种动态路由协议（$($protos -join '/')），超出常见业务所需，请逐条核对业务归属。" `
                "第5章" "按业务最小化启用路由协议，关闭无用途协议与闲置网段。"
        } else {
            Add-Result "5.4" "网络安全" "边界互联节点与路由最小化" "pass" `
                "动态路由协议共 $protoCnt 种，未见明显冗余；路由业务归属仍需人工逐条确认。" `
                "第5章" "定期清理无业务归属路由。"
        }
    } else {
        Add-Result "5.4" "网络安全" "边界互联节点与路由最小化" "manual" `
            "未采集到路由表回显，无法自动判定。" "第5章" "补采 display ip routing-table / show ip route。"
    }
    # ---- 5.5 安全区域划分 ----
    if (NonEmpty $vlan) {
        $vlanCnt2 = Count-Vlans $vlan
        if ($vlanCnt2 -ge 3) {
            Add-Result "5.5" "网络安全" "按业务性质划分安全区域" "pass" `
                "发现 $vlanCnt2 个 VLAN，具备区域划分基础。请比对安全区域划分图确认服务器/存储/嵌入式设备分区。" `
                "第5章" "区域边界应配套隔离控制（ACL/防火墙）。"
        } else {
            Add-Result "5.5" "网络安全" "按业务性质划分安全区域" "manual" `
                "VLAN 数量偏少（$vlanCnt2），是否满足分区要求需结合区域划分图人工确认。" "第5章" "按用途划分安全区域。"
        }
    } else {
        Add-Result "5.5" "网络安全" "按业务性质划分安全区域" "manual" "未采集到 VLAN 回显。" "第5章" "补采 display vlan。"
    }
    # ---- 5.6 设备安全配置最小化 ----
    if (NonEmpty $cfg) {
        $telnetOn = Count-Lines $cfg 'telnet server enable|telnet enable' $true
        $ftpOn = Count-Lines $cfg 'ftp server enable' $true
        if ($telnetOn -gt 0 -or $ftpOn -gt 0) {
            Add-Result "5.6" "网络安全" "网络设备最小化安全配置" "fail" `
                "配置中发现 telnet 服务启用 $telnetOn 处、ftp 服务启用 $ftpOn 处，存在非必要服务；请同时核对账户清单与 ACL 最小放行。" `
                "第5章" "关闭 telnet/ftp 等非必要服务，账户按最小化授权。"
        } else {
            Add-Result "5.6" "网络安全" "网络设备最小化安全配置" "pass" `
                "未发现 telnet/ftp 服务启用。账户与 ACL 的最小化仍需按指导书 5.6 方法（4）（5）人工核对。" `
                "第5章" "定期清理冗余账户与全放行策略。"
        }
    } else {
        Add-Result "5.6" "网络安全" "网络设备最小化安全配置" "manual" "未采集到设备配置回显。" "第5章" "补采 display current-configuration / show running-config。"
    }
    # ---- 5.7 唯一管理服务 ----
    if (NonEmpty $sshSt) {
        $telnetCfg = Count-Lines $cfg 'telnet server enable|telnet enable' $true
        if ($sshSt -imatch 'enable|active|start') {
            if ($telnetCfg -gt 0) {
                Add-Result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "fail" `
                    "SSH 已启用，但配置中 telnet 仍开启（$telnetCfg 处），管理服务不唯一；管理终端 ACL 限制请按 VTY 配置人工核对。" `
                    "第5章" "仅保留 SSH，关闭 Telnet，并以 ACL 限定管理终端。"
            } else {
                Add-Result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "pass" `
                    "SSH 管理服务启用且未见 telnet。管理终端来源限制（VTY ACL）需人工核对（见指导书 5.7 方法（3））。" `
                    "第5章" "VTY 线路绑定 ACL 仅放行管理网段。"
            }
        } else {
            Add-Result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "fail" `
                "SSH 服务状态回显未显示启用，且存在明文管理风险。请核实管理方式。" "第5章" "启用 SSH 并关闭明文管理服务。"
        }
    } else {
        Add-Result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "manual" "未采集到 SSH 服务状态回显。" "第5章" "补采 display ssh server status / show ip ssh。"
    }
    # ---- 5.8 加密管理 ----
    if (NonEmpty $sshSt) {
        $v2 = Count-Lines $sshSt 'version 2|2\.0|SSH2' $true
        $telnetCfg2 = Count-Lines $cfg 'telnet server enable|telnet enable' $true
        $httpOn = Count-Lines $cfg 'http server enable|web-manager http' $true
        if ($v2 -gt 0 -and $telnetCfg2 -eq 0 -and $httpOn -eq 0) {
            Add-Result "5.8" "网络安全" "远程管理采用加密通道" "pass" `
                "SSH 版本为 2.0，未发现 telnet/http 明文管理。防火墙设备 HTTPS 证书请人工核对。" `
                "第5章" "设备与防护设备管理均应走 SSH/HTTPS。"
        } else {
            Add-Result "5.8" "网络安全" "远程管理采用加密通道" "fail" `
                "加密管理不达标：SSH2 特征 $v2 处、telnet $telnetCfg2 处、http 管理 $httpOn 处。存在明文管理风险。" `
                "第5章" "关闭 Telnet/HTTP 管理，统一 SSH2/HTTPS。"
        }
    } else {
        Add-Result "5.8" "网络安全" "远程管理采用加密通道" "manual" "未采集到 SSH 状态回显。" "第5章" "补采 SSH 状态并核对加密层次。"
    }
    # ---- 5.9 异构防护（台账类） ----
    $brands = @([regex]::Matches($ver, 'HUAWEI|Huawei|H3C|Comware|Ruijie|RG[- ]?[A-Z0-9]+|S[0-9]{4}|USG[0-9]+|SecPath') | ForEach-Object { $_.Value } | Select-Object -Unique)
    Add-Result "5.9" "网络安全" "同功能防护设备异构部署" "manual" `
        "需设备台账比对：同链路承担相同功能的设备应不同品牌/架构。可参考已采集的 display version 型号信息：$($brands -join ' ')" `
        "第5章" "关键链路同功能设备应异构，防共因失效。"
    # ---- 5.10 组播控制 ----
    $mcastNoEntry = Count-Lines $mcast 'total 0 entry|0 entry altogether|no (multicast )?route|无(组播)?(表项|路由)|^\(空' $true
    if (-not (NonEmpty $mcast) -or $mcastNoEntry -gt 0) {
        Add-Result "5.10" "网络安全" "组播源/地址/成员控制" "pass" `
            "未发现组播路由表项（无表项或未启用组播业务）。" "第5章" "后续启用组播须先登记并配置边界过滤。"
    } else {
        Add-Result "5.10" "网络安全" "组播源/地址/成员控制" "manual" `
            "存在组播路由表项，需按登记清单核对组播源与成员控制（见指导书 5.10 方法（4）边界过滤核查）。" `
            "第5章" "组播边界应配置过滤，未登记组播应清除。"
    }
    # ---- 5.11 设备备份 ----
    if (NonEmpty $hrp) {
        Add-Result "5.11" "网络安全" "重要网络与防护设备备份" "pass" `
            "采集到 HRP（双机热备）状态回显，具备热备机制。冷备设备与配置备份周期仍需按台账人工核对。" `
            "第5章" "备份设备定期启用验证，配置按期备份。"
    } else {
        Add-Result "5.11" "网络安全" "重要网络与防护设备备份" "manual" `
            "未采集到双机热备状态（或该设备无热备）。请按指导书 5.11 核对备份制度、冷备与配置备份存档。" `
            "第5章" "重要设备应有备份并演练恢复。"
    }
    # ---- 5.12 远程传输两层加密 ----
    if (NonEmpty $ike) {
        Add-Result "5.12" "网络安全" "远程传输两层加密保护" "manual" `
            "采集到 IKE SA（IPSec 隧道协商）信息，网络层加密在位。链路层/信源层加密层次需按方案人工确认（见指导书 5.12）。" `
            "第5章" "远程传输应满足两层加密（链路/网络/信源任两层）。"
    } else {
        Add-Result "5.12" "网络安全" "远程传输两层加密保护" "manual" `
            "未采集到 IKE SA 回显（无 IPSec 隧道或未采集）。按传输加密方案人工核查层次组合。" `
            "第5章" "补采 display ike sa 或核对加密机部署。"
    }
    # ---- 5.13 管理中心（管理类） ----
    Add-Result "5.13" "网络安全" "网络安全管理中心与指定管理终端" "manual" `
        "管理类核查：实地查看管理中心与专用终端，登录管理平台核对纳管清单与终端 IP（见指导书 5.13）。" `
        "第5章" "多类设备应经管理中心统一单独管理。"
    # ---- 5.14 等级隔离 ----
    $sessText = Get-Sec 'display firewall session table'
    $sessNote = "未采集到防火墙会话表"
    if (NonEmpty $sessText) { $sessNote = "已采集到会话表，可比对未审批通道" }
    Add-Result "5.14" "网络安全" "与外部网络按等级物理/逻辑隔离" "manual" `
        "隔离类核查：核对隔离拓扑、实地查看物理隔离点；防火墙会话表（display firewall session table）可用于比对跨域通道，本次$sessNote。" `
        "第5章" "隔离方式须与防护等级对应。"
    # ---- 5.15 边界防护与告警审计 ----
    Add-Result "5.15" "网络安全" "边界防护设备的交换控制/告警/审计/阻断能力" "manual" `
        "能力类核查：登录防火墙/网闸查看白名单策略、告警配置与审计日志（见指导书 5.15 方法（2）-（5））。" `
        "第5章" "边界应具备实时告警、审计、阻断与定位能力。"
    # ---- 5.16 细粒度访问控制 ----
    if (NonEmpty $acl) {
        $aclRules = Count-Lines $acl '^ *(rule|Rule|permit|deny) ' $false
        $portRules = Count-Lines $acl 'eq (80|443|22|23|3389|[0-9]+)|destination-port|port (eq|range)' $true
        if ($aclRules -ge 3) {
            Add-Result "5.16" "网络安全" "全网细粒度访问控制" "pass" `
                "ACL 规则共 $aclRules 条，其中端口级规则特征 $portRules 处，具备细粒度基础。五元组颗粒度与区域间覆盖需人工抽查（见指导书 5.16 方法（4）（5））。" `
                "第5章" "区域间与主机间均应部署五元组级白名单策略。"
        } else {
            Add-Result "5.16" "网络安全" "全网细粒度访问控制" "fail" `
                "ACL 规则仅 $aclRules 条，颗粒度或覆盖不足，疑似粗放策略或未部署访问控制。" `
                "第5章" "按最小放行原则细化区域间访问控制。"
        }
    } else {
        Add-Result "5.16" "网络安全" "全网细粒度访问控制" "manual" "未采集到 ACL 回显。" "第5章" "补采 display acl all / show access-lists。"
    }
    # ---- 5.17 租用线路三层加密 ----
    if (NonEmpty $ike) {
        Add-Result "5.17" "网络安全" "租用线路三层加密保护" "manual" `
            "采集到 IKE SA，网络层加密在位。链路层与信源层加密、密钥管理记录需人工核查（见指导书 5.17 方法（5））。" `
            "第5章" "租线传输应满足三层加密。"
    } else {
        Add-Result "5.17" "网络安全" "租用线路三层加密保护" "manual" `
            "未采集到 IKE SA。租用线路如存在，请按三层加密方案核查。" "第5章" "补采 display ike sa。"
    }
    # ---- 5.18 接入认证 ----
    if (NonEmpty $dot1x) {
        if ($dot1x -imatch 'enable|enabled|active|已启用') {
            Add-Result "5.18" "网络安全" "802.1x 接入认证" "pass" `
                "802.1x 认证已启用。认证要素（端口+IP+MAC 绑定）需登录 Radius/准入平台人工核对（见指导书 5.18 方法（4））。" `
                "第5章" "认证要素至少含端口、IP、MAC。"
        } else {
            $firstLine = (($dot1x -split "`r?`n") | Select-Object -First 1)
            if ($null -ne $firstLine -and $firstLine.Length -gt 40) { $firstLine = $firstLine.Substring(0, 40) }
            Add-Result "5.18" "网络安全" "802.1x 接入认证" "fail" `
                "802.1x 状态回显未显示启用（回显首行：$firstLine）。" "第5章" "启用 802.1x 或同强度接入认证。"
        }
    } else {
        Add-Result "5.18" "网络安全" "802.1x 接入认证" "manual" "未采集到 dot1x 回显。" "第5章" "补采 display dot1x / show dot1x summary。"
    }
    # ---- 5.19 终端逻辑隔离 ----
    if (NonEmpty $pisol) {
        $pgroups = Count-Lines $pisol 'group|Group' $false
        Add-Result "5.19" "网络安全" "用户计算机间逻辑隔离" "pass" `
            "发现端口隔离组特征 $pgroups 处，接入层已部署终端隔离。主机防火墙状态需在终端侧另行核查。" `
            "第5章" "终端间应默认不可互访，互访经集中服务。"
    } else {
        Add-Result "5.19" "网络安全" "用户计算机间逻辑隔离" "manual" `
            "未采集到端口隔离组回显（可能采用 PVLAN 等其他隔离方式）。请按指导书 5.19 方法（4）在终端侧核查防火墙，并做互访测试。" `
            "第5章" "接入层应部署端口隔离或等效措施。"
    }
    # ---- 5.20 全网行为审计 ----
    Add-Result "5.20" "网络安全" "全网行为审计且日志留存≥180天" "manual" `
        "平台类核查：登录行为审计/上网管理系统核对覆盖范围、最早可查日志时间与要素完整性（见指导书 5.20 方法（2）-（5））。" `
        "第5章" "审计日志留存不少于 180 天且防删改。"
    # ---- 5.21 攻击实时监视 ----
    Add-Result "5.21" "网络安全" "攻击/违规行为实时监视告警阻断" "manual" `
        "平台类核查：查看 IDS/IPS/态势感知的策略库、实时告警与处置闭环记录（见指导书 5.21 方法（2）-（5））。" `
        "第5章" "应能实时告警、阻断并定位攻击源。"
    # ---- 5.22 存储管理网隔离 ----
    if (NonEmpty $vlan) {
        Add-Result "5.22" "网络安全" "数据存储系统管理网与应用网隔离" "manual" `
            "VLAN 回显可用于核对存储管理口/业务口分属不同 VLAN（见指导书 5.22 方法（3）（4））；请结合存储设备台账与连线照片确认。" `
            "第5章" "存储管理流量与业务流量分网。"
    } else {
        Add-Result "5.22" "网络安全" "数据存储系统管理网与应用网隔离" "manual" "未采集到 VLAN 回显。" "第5章" "补采 display vlan 并核对存储端口归属。"
    }
    # ---- 5.23 设备配备合理必要性 ----
    if (NonEmpty $brief) {
        $ifLines = Get-IfLines $brief
        $totalIf = $ifLines.Count
        $downIf  = @($ifLines | Where-Object { $_ -imatch 'down|notconnect' }).Count
        if ($totalIf -gt 0) {
            $pctDown = [math]::Floor($downIf * 100 / $totalIf)
            if ($pctDown -ge 50) {
                Add-Result "5.23" "网络安全" "网络设备配备合理必要性" "manual" `
                    "端口状态：总 $totalIf 口中 $downIf 口 DOWN（$pctDown%）。闲置比例偏高，请按台账逐台甄别必要冗余与无用途设备。" `
                    "第5章" "清退无用途设备与闲置链路。"
            } else {
                Add-Result "5.23" "网络安全" "网络设备配备合理必要性" "pass" `
                    "端口状态：总 $totalIf 口中 $downIf 口 DOWN（$pctDown%），未见大面积闲置。设备必要性仍需台账逐台核对。" `
                    "第5章" "定期清点设备用途。"
            }
        } else {
            Add-Result "5.23" "网络安全" "网络设备配备合理必要性" "manual" "端口状态回显解析不出接口行，请人工核对台账。" "第5章" "定期清点设备用途。"
        }
    } else {
        Add-Result "5.23" "网络安全" "网络设备配备合理必要性" "manual" "未采集到端口状态回显。" "第5章" "补采 display interface brief / show interfaces status。"
    }

    Write-Host ""
    Write-Host "核查完成，共 $($script:R.Count) 项（网络设备 $($script:DevCount) 台：$($script:DevNames)）"
    Generate-Html
    Generate-Xls
}

# ---------- 报告生成（与组件脚本同款模板） ----------
function Json-Esc([string]$s) {
    if ($null -eq $s) { return "" }
    $s = $s -replace '\\', '\\\\'
    $s = $s -replace '"', '\"'
    $s = $s -replace "`r?`n", ' '
    $s = $s -replace "`t", ' '
    return $s
}

$REPORT_TAG = "网络设备"
$REPORT_TITLE = "网络设备配置核查报告"

function Generate-Html {
    $outfile = Join-Path $OutDir "配置核查报告_${REPORT_TAG}_${Stamp}.html"
    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $pass = @($script:R | Where-Object { $_.status -eq 'pass' }).Count
    $fail = @($script:R | Where-Object { $_.status -eq 'fail' }).Count
    $manual = @($script:R | Where-Object { $_.status -eq 'manual' }).Count
    $na = @($script:R | Where-Object { $_.status -eq 'na' }).Count
    $total = $script:R.Count
    $ppass = "0.0"; $pfail = "0.0"; $pmanual = "0.0"; $pna = "0.0"
    if ($total -gt 0) {
        $ppass = "{0:F1}" -f ($pass * 100 / $total)
        $pfail = "{0:F1}" -f ($fail * 100 / $total)
        $pmanual = "{0:F1}" -f ($manual * 100 / $total)
        $pna = "{0:F1}" -f ($na * 100 / $total)
    }
    $dataRows = ($script:R | ForEach-Object {
        ('{"ch":"' + (Json-Esc $_.chapter) + '","id":"' + (Json-Esc $_.id) + '","cat":"' + (Json-Esc $_.cat) +
         '","title":"' + (Json-Esc $_.title) + '","status":"' + $_.status + '","detail":"' + (Json-Esc $_.detail) +
         '","rec":"' + (Json-Esc $_.rec) + '","guide":"' + (Json-Esc $_.guide) + '"}')
    }) -join ","
    $dataJson = "[" + $dataRows + "]"
    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $html = @'
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="GBK">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root{--bg:#f5f7fa;--card:#fff;--ink:#1f2937;--muted:#6b7280;--line:#e5e7eb;--brand:#0f3057;--brand2:#1a3c6e;--pass:#15803d;--pass-bg:#ecfdf5;--pass-br:#bbf7d0;--fail:#b91c1c;--fail-bg:#fef2f2;--fail-br:#fecaca;--manual:#b45309;--manual-bg:#fffbeb;--manual-br:#fde68a;--na:#4b5563;--na-bg:#f3f4f6;--na-br:#e5e7eb;}
*{box-sizing:border-box;}
body{margin:0;font:14px/1.65 "Segoe UI","Microsoft YaHei",system-ui,sans-serif;color:var(--ink);background:var(--bg);}
.wrap{max-width:1240px;margin:0 auto;padding:24px 20px 60px;}
header{background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 60%,#2563eb 130%);color:#fff;border-radius:12px;padding:26px 30px;margin-bottom:20px;}
header h1{margin:0 0 6px;font-size:22px;letter-spacing:.5px;}
header .sub{opacity:.85;font-size:13px;}
.meta{display:flex;flex-wrap:wrap;gap:8px 28px;margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.25);font-size:13px;}
.meta div{opacity:.95;}
.dash{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:20px;}
.stat{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:18px 16px 14px;position:relative;overflow:hidden;cursor:pointer;transition:transform .15s, box-shadow .15s;}
.stat:hover{transform:translateY(-2px);box-shadow:0 6px 18px rgba(15,48,87,.10);}
.stat .num{font-size:34px;font-weight:700;line-height:1.1;font-variant-numeric:tabular-nums;}
.stat .lbl{color:var(--muted);font-size:13px;margin-top:2px;}
.stat .bar{height:4px;border-radius:2px;margin-top:12px;background:var(--line);}
.stat .bar i{display:block;height:100%;border-radius:2px;}
.stat.s-pass .num{color:var(--pass);} .stat.s-pass .bar i{background:var(--pass);}
.stat.s-fail .num{color:var(--fail);} .stat.s-fail .bar i{background:var(--fail);}
.stat.s-manual .num{color:var(--manual);} .stat.s-manual .bar i{background:var(--manual);}
.stat.s-na .num{color:var(--na);} .stat.s-na .bar i{background:var(--na);}
.stat .pct{position:absolute;right:14px;top:16px;font-size:12px;color:var(--muted);}
.toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:16px;}
.toolbar input[type=text]{flex:1;min-width:200px;padding:9px 14px;border:1px solid var(--line);border-radius:8px;font-size:13px;outline:none;background:var(--card);}
.toolbar input[type=text]:focus{border-color:#2563eb;box-shadow:0 0 0 3px rgba(37,99,235,.12);}
.filters{display:flex;gap:6px;flex-wrap:wrap;}
.fbtn{border:1px solid var(--line);background:var(--card);color:var(--ink);padding:7px 14px;border-radius:20px;font-size:13px;cursor:pointer;transition:all .15s;}
.fbtn:hover{border-color:#2563eb;color:#2563eb;}
.fbtn.on{background:var(--brand);border-color:var(--brand);color:#fff;}
.count{color:var(--muted);font-size:12px;margin-left:4px;}
.panel{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden;}
table{border-collapse:collapse;width:100%;font-size:13px;}
thead th{background:#f8fafc;color:#334155;text-align:left;font-weight:600;padding:10px 12px;border-bottom:2px solid var(--line);white-space:nowrap;position:sticky;top:0;z-index:5;}
tbody td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top;}
tbody tr:hover{background:#f8fafc;}
td.id{font-family:Consolas,monospace;font-weight:600;white-space:nowrap;}
td.cat{white-space:nowrap;color:var(--muted);}
td.title{min-width:180px;}
.badge{display:inline-block;padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600;white-space:nowrap;border:1px solid;}
.badge-pass{color:var(--pass);background:var(--pass-bg);border-color:var(--pass-br);}
.badge-fail{color:var(--fail);background:var(--fail-bg);border-color:var(--fail-br);}
.badge-manual{color:var(--manual);background:var(--manual-bg);border-color:var(--manual-br);}
.badge-na{color:var(--na);background:var(--na-bg);border-color:var(--na-br);}
td.detail,td.rec{color:#374151;max-width:320px;}
.guide{color:var(--muted);font-size:12px;max-width:260px;}
.empty{padding:60px;text-align:center;color:var(--muted);}
footer{margin-top:26px;color:var(--muted);font-size:12px;text-align:center;}
@media(max-width:900px){.dash{grid-template-columns:repeat(2,1fr);}}
@media print{.toolbar{display:none;} .panel{border:none;} body{background:#fff;}}
</style>
</head>
<body>
<div class="wrap">
<header>
  <h1>__TITLE__</h1>
  <div class="sub">参考标准：配置核查作业指导书v2.2　|　核查方式：设备命令回显解析（第5章 网络安全）</div>
  <div class="meta">
    <div>设备：__DEVCOUNT__ 台（__DEVNAMES__）</div>
    <div>核查时间：__TIME__</div>
  </div>
</header>
<div class="dash">
  <div class="stat s-pass" onclick="fset('pass')"><div class="num">__PASS__</div><div class="lbl">合规</div><div class="bar"><i style="width:__PPASS__%"></i></div><div class="pct">__PPASS__%</div></div>
  <div class="stat s-fail" onclick="fset('fail')"><div class="num">__FAIL__</div><div class="lbl">不合规</div><div class="bar"><i style="width:__PFAIL__%"></i></div><div class="pct">__PFAIL__%</div></div>
  <div class="stat s-manual" onclick="fset('manual')"><div class="num">__MANUAL__</div><div class="lbl">需人工核查</div><div class="bar"><i style="width:__PMANUAL__%"></i></div><div class="pct">__PMANUAL__%</div></div>
  <div class="stat s-na" onclick="fset('na')"><div class="num">__NA__</div><div class="lbl">不适用</div><div class="bar"><i style="width:__PNA__%"></i></div><div class="pct">__PNA__%</div></div>
</div>
<div class="toolbar">
  <input id="q" type="text" placeholder="搜索编号 / 核查项 / 详情…">
  <div class="filters">
    <button class="fbtn on" data-f="all">全部<span class="count">__TOTAL__</span></button>
    <button class="fbtn" data-f="fail">不合规<span class="count">__FAIL__</span></button>
    <button class="fbtn" data-f="manual">需人工<span class="count">__MANUAL__</span></button>
    <button class="fbtn" data-f="pass">合规<span class="count">__PASS__</span></button>
    <button class="fbtn" data-f="na">不适用<span class="count">__NA__</span></button>
  </div>
</div>
<div class="panel">
<table id="tbl">
<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>
<tbody id="tb"></tbody>
</table>
<div class="empty" id="empty" style="display:none">没有匹配的核查项</div>
</div>
<footer>本报告由配置核查工具自动生成 · __TIME__</footer>
</div>
<script>
var DATA =
__DATA__
;
var ST = {pass:"合规", fail:"不合规", manual:"需人工核查", na:"不适用"};
var curF = "all";
function esc(s){var d=document.createElement("div");d.textContent=s==null?"":s;return d.innerHTML;}
function render(){
  var q = document.getElementById("q").value.trim().toLowerCase();
  var tb = document.getElementById("tb"); tb.innerHTML = "";
  var n = 0;
  DATA.forEach(function(x){
    if(curF!="all" && x.status!=curF) return;
    if(q && (x.id+" "+x.title+" "+x.detail+" "+x.cat+" "+x.ch).toLowerCase().indexOf(q)<0) return;
    n++;
    var tr = document.createElement("tr");
    tr.innerHTML = "<td>"+esc(x.ch)+"</td><td class='id'>"+esc(x.id)+"</td><td class='cat'>"+esc(x.cat)+"</td>"+
      "<td class='title'>"+esc(x.title)+"</td>"+
      "<td><span class='badge badge-"+x.status+"'>"+ST[x.status]+"</span></td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+esc(x.rec)+"</td>"+
      "<td class='guide'>"+esc(x.guide)+"</td>";
    tb.appendChild(tr);
  });
  document.getElementById("empty").style.display = n? "none":"block";
}
function fset(f){
  curF = f;
  var bs = document.querySelectorAll(".fbtn");
  for(var i=0;i<bs.length;i++){ bs[i].className = "fbtn" + (bs[i].getAttribute("data-f")==f ? " on" : ""); }
  render();
}
(function(){
  var bs = document.querySelectorAll(".fbtn");
  for(var i=0;i<bs.length;i++){ bs[i].onclick = (function(b){ return function(){ fset(b.getAttribute("data-f")); }; })(bs[i]); }
})();
document.getElementById("q").addEventListener("input", render);
render();
</script>
</body>
</html>
'@
    $html = $html.Replace('__TITLE__', (Html-Esc $REPORT_TITLE)).Replace('__DEVCOUNT__', (Html-Esc "$($script:DevCount)")).Replace('__DEVNAMES__', (Html-Esc $script:DevNames))
    $html = $html.Replace('__TIME__', $now).Replace('__DATA__', $dataJson).Replace('__TOTAL__', "$total")
    $html = $html.Replace('__PASS__', "$pass").Replace('__FAIL__', "$fail").Replace('__MANUAL__', "$manual").Replace('__NA__', "$na")
    $html = $html.Replace('__PPASS__', $ppass).Replace('__PFAIL__', $pfail).Replace('__PMANUAL__', $pmanual).Replace('__PNA__', $pna)
    # 与其余 win 脚本（VBS/ANSI）一致：GBK 输出；merge_report.vbs 按 ANSI 读取汇总
    $gbk = [System.Text.Encoding]::GetEncoding(936)
    [System.IO.File]::WriteAllText($outfile, $html, $gbk)
    Write-Host "HTML报告已生成：$outfile"
}

function Generate-Xls {
    $outfile = Join-Path $OutDir "配置核查报告_${REPORT_TAG}_${Stamp}.xls"
    $pass = @($script:R | Where-Object { $_.status -eq 'pass' }).Count
    $fail = @($script:R | Where-Object { $_.status -eq 'fail' }).Count
    $manual = @($script:R | Where-Object { $_.status -eq 'manual' }).Count
    $na = @($script:R | Where-Object { $_.status -eq 'na' }).Count
    $now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">')
    [void]$sb.AppendLine('<head><meta charset="GBK">')
    [void]$sb.AppendLine('<!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>配置核查报告</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->')
    [void]$sb.AppendLine('<style>table{border-collapse:collapse;}th,td{border:1px solid #999;padding:4px 6px;font-family:"Microsoft YaHei",Arial;font-size:12px;mso-number-format:"\@";}th{background:#1a3c6e;color:#fff;font-weight:bold;}</style>')
    [void]$sb.AppendLine('</head><body>')
    [void]$sb.AppendLine("<p><b>$REPORT_TITLE</b>　设备：$(Html-Esc $script:DevNames)　核查时间：$now</p>")
    [void]$sb.AppendLine('<p>参考标准：配置核查作业指导书v2.2 第5章 网络安全</p>')
    [void]$sb.AppendLine("<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>")
    [void]$sb.AppendLine('<table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>')
    foreach ($r in $script:R) {
        $scn = "不适用"
        switch ($r.status) {
            "pass"   { $scn = "合规" }
            "fail"   { $scn = "不合规" }
            "manual" { $scn = "需人工核查" }
        }
        [void]$sb.AppendLine("<tr><td>$(Html-Esc $r.chapter)</td><td>$(Html-Esc $r.id)</td><td>$(Html-Esc $r.cat)</td><td>$(Html-Esc $r.title)</td><td>$scn</td><td>$(Html-Esc $r.detail)</td><td>$(Html-Esc $r.rec)</td><td>$(Html-Esc $r.guide)</td></tr>")
    }
    [void]$sb.AppendLine('</table></body></html>')
    $gbk = [System.Text.Encoding]::GetEncoding(936)
    [System.IO.File]::WriteAllText($outfile, $sb.ToString(), $gbk)
    Write-Host "Excel(.xls)报告已生成：$outfile"
}

# ---------- 入口 ----------
switch ($Command.ToLower()) {
    "init"  { Do-Init }
    "check" { Do-Check }
    default {
        Write-Host "用法：powershell -ExecutionPolicy Bypass -File check_network.ps1 [init|check]"
        Write-Host "  init  生成采集工作区与三厂商命令清单（首次使用）"
        Write-Host "  check 解析已采集的设备回显并生成核查报告"
        Write-Host "  （或直接双击 check_network.bat init / check）"
    }
}
