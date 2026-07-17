#!/bin/bash
# ============================================================
# 配置核查工具 - 中标麒麟/银河麒麟版（无需Python，纯Bash实现）
# 适用系统：中标麒麟 NeoKylin（yum/rpm，多基于CentOS/RHEL）
#           银河麒麟 Kylin OS（apt/dpkg，多基于Ubuntu/Debian）
# 参考标准：配置核查作业指导书正式版2026_4_1
# 运行方式：sudo bash run.sh 或 sudo bash check_kylin.sh
# 输出文件：output/配置核查报告_日期时间.html /.xls
# ============================================================
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT_DIR="$SCRIPT_DIR/output"
mkdir -p "$OUT_DIR"
STAMP="$(date +%Y%m%d_%H%M%S)"

# ---------- 结果存储（并行数组）----------
R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=()
R_COUNT=0

# 《配置核查作业指导书正式版2026_4_1》目录页码与条款原文对照表（按核查编号查询）
# 用于在报告中提示测试人员应翻阅指导书哪一页、对照哪一条原文表述
guide_ref() {
    local page="" title=""
    case "$1" in
        1.1) page=1; title="操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.2) page=6; title="操作系统应安装防病毒软件并及时升级" ;;
        1.3) page=8; title="操作系统应按需求裁剪服务和端口" ;;
        1.4) page=11; title="操作系统应具备防火墙功能" ;;
        1.5) page=12; title="操作系统应停用冗余网络设置" ;;
        1.6) page=14; title="操作系统远程管理应开放唯一管理服务，指定管理终端并采取传输加密保护措施" ;;
        1.7) page=19; title="数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" ;;
        1.8) page=26; title="数据库管理系统应删除冗余存储过程" ;;
        1.9) page=32; title="数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" ;;
        1.10) page=37; title="数据库管理系统应具有自主访问控制功能" ;;
        1.11) page=46; title="数据库管理系统应具有备份和恢复功能" ;;
        1.12) page=53; title="数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.13) page=60; title="数据库管理系统的数据应和其它应用的数据分类独立存储" ;;
        1.14) page=64; title="中间件应采取限制运行权限和使用安全管理通道等安全加固措施" ;;
        1.15) page=70; title="应具备数据库管理系统超级管理员远程登录限制能力" ;;
        1.16) page=76; title="应具备数据库管理系统输入（参数）检查能力" ;;
        1.16-17) page="76、84"; title="应具备数据库管理系统输入（参数）检查能力 / 应限制操作系统开放的远程管理服务或端口" ;;
        1.17) page=84; title="应限制操作系统开放的远程管理服务或端口" ;;
        1.18) page=84; title="应限制用户对服务器资源的最大或最小使用限度" ;;
        1.19) page=85; title="应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.20) page=88; title="数据库管理系统应配置安全策略" ;;
        1.21) page=97; title="数据库管理系统应具有行级或列级审计功能" ;;
        1.22) page=101; title="数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) page=108; title="数据库管理系统仅为应用服务器提供访问服务" ;;
        1.24) page=113; title="应具备日志审计能力，审计日志至少保留180天" ;;
        1.25) page=121; title="检查是否具备边界保护能力，是否可以抗攻击、防篡改" ;;
        1.26|1.26_m) page=128; title="检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" ;;
        2.1) page=133; title="服务器和用户计算机应设置登录口令" ;;
        2.2) page=136; title="用户计算机应根据需要安装补丁程序" ;;
        2.3) page=139; title="用户应设置用户应用口令，通过认证后使用信息服务" ;;
        2.4) page=144; title="用户计算机应具有互不相同的用户名和口令" ;;
        2.5) page=147; title="用户计算机应关闭冗余系统服务和端口" ;;
        2.6) page=151; title="用户计算机应具备阻断和告警非法连接互联网的能力" ;;
        2.7) page=155; title="用户计算机应具有文件保护功能" ;;
        2.8) page=159; title="应采取终端管控措施（统一配置、安全加固、网络访问控制、外设接口管控、软件进程管控、无线模块禁用等）" ;;
        2.9) page=162; title="用户计算机应禁止安装与工作无关的软件" ;;
        2.10) page=164; title="用户计算机USB接口应禁止私自连接对拷线和手机、媒体播放设备等个人移动电子设备" ;;
        2.11|2.11_m) page=167; title="用户计算机登录应使用基于专用物理部件或生物特征的多因素身份认证方式，应设置超时锁屏（不超过5min），口令不少于10个字符，更换周期不超过30d" ;;
        2.12) page=168; title="用户计算机应物理拆除Wi-Fi、红外、蓝牙等无线模块" ;;
        2.13) page=171; title="用户计算机应采取非法外联阻断、文件输出管控等措施" ;;
        2.14) page=174; title="应具备用户行为审计能力；审计日志应至少保留xx天" ;;
        2.15) page=177; title="检查安全策略配置情况和设置功能" ;;
        2.16) page=178; title="检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" ;;
        3.1) page=183; title="集中存储的涉密数据应采取加密保护措施" ;;
        3.2) page=186; title="用户计算机之间的数据交互、文件传输等应统一管控和审计" ;;
        3.3) page=189; title="涉密存储载体在降密级使用前或重大演训活动结束后，应采取数据写覆盖方法及时清除数据" ;;
        3.4) page=191; title="对确定销毁的涉密载体，应采取消磁、粉碎、溶解、化浆和熔化等方法进行销毁" ;;
        3.5) page=191; title="网络、系统、应用和用户行为等日志应采取读写控制、加密、变换、完整性校验等保护措施" ;;
        3.6) page=195; title="网络边界应具备信息过滤、敏感内容识别等数据防泄漏能力" ;;
        3.7) page=202; title="数据存储系统管理登录至少采取验证码等增强措施，修改默认设置，审计日志至少保留180天" ;;
        3.8) page=207; title="数据存储系统应根据重要程度划分不同存储区块，并设置用户访问权限" ;;
        3.9) page=213; title="检查数据传输过程是否按要求加密，传输路径是否合理，是否统一管控、留有日志记录、具有防泄漏措施" ;;
        3.10) page=222; title="检查数据共享是否合理，是否存在安全隐患" ;;
        3.11) page=228; title="检查对数据的访问是否按照权限分级访问，是否对访问行为进行审计" ;;
        3.12) page=233; title="检查数据采集是否超出业务需求范围" ;;
        3.13) page=235; title="检查数据各环节处理是否满足密级相应的保密要求" ;;
        3.14) page=239; title="检查对数据的访问是否按照权限分级访问，对大数据的访问是否提供统一管控和访问控制" ;;
        4.1) page=246; title="应用系统软件应采取备份措施" ;;
        4.2) page=248; title="应用系统软件应及时安装补丁程序，且更新所用升级包应经过安全性测试" ;;
        4.3) page=249; title="基于可信根对应用系统软件进行可信验证，可信性受到破坏后报警" ;;
        4.4) page=251; title="提供公共信息服务的服务器应与涉密信息服务器分设，专用服务器应只提供专用服务" ;;
        4.5) page=254; title="提供公共信息服务的服务器应具备防DDoS攻击能力" ;;
        4.6) page=257; title="Web应用系统服务器应采取Web防护措施" ;;
        4.7) page=260; title="网站服务宜以静态页面形式发布" ;;
        4.8) page=262; title="网站应采取网页防篡改措施，防止对信息内容的非法修改" ;;
        4.9) page=263; title="Web应用系统应具备防范SQL注入、跨站脚本等攻击能力" ;;
        4.10) page=265; title="Web应用系统应具备执行代码有效验证能力" ;;
        4.11) page=267; title="应具备用户授权访问控制能力" ;;
        4.12) page=269; title="应具备访问应用最大并发会话连接数限制能力" ;;
        4.13) page=271; title="业务管理终端专设专用" ;;
        4.14) page=272; title="应具有基于用户角色的授权访问控制能力，主体细粒度达到用户级或进程级，客体细粒度达到文件级、表和记录级、字段级" ;;
        4.15) page=275; title="文电等文档专用业务处理系统应具有签名验证、密级标识等功能" ;;
        4.16) page=278; title="远程管理应采取加密保护措施" ;;
        4.17) page=280; title="应支持管理员、安全员、审计员三权分立的职责划分，禁止设立超级管理员" ;;
        4.18) page=284; title="业务管理终端登录应采取网络地址限制措施" ;;
        4.19) page=287; title="应具有结束会话、限定登录错误次数和自动退出等登录失败处理功能" ;;
        4.20) page=289; title="宜使用自主设计开发的网络服务、协议、接口等，增强应用安全" ;;
        4.21) page=292; title="应具有基于专用物理部件或生物特征多因素的、与JD密码算法相结合的数字证书用户身份认证功能" ;;
        4.22) page=294; title="应更改Web应用系统默认服务发布端口" ;;
        4.23) page=298; title="应分开设置管理端口与应用端口" ;;
        4.24) page=300; title="应用服务和数据存储应部署在不同的服务器上" ;;
        4.25) page=304; title="应具有对所有访问行为和管理行为的日志审计功能，审计日志应至少保留180d" ;;
        4.26) page=309; title="应经过代码级安全漏洞挖掘" ;;
        4.27) page=311; title="检查是否具备对人机接口输入、网络通信输入、文件输入的数据进行格式和长度检查的功能" ;;
        4.28) page=313; title="检查是否能够有效检测并防御SQL注入、网页篡改、跨站脚本、拒绝服务等应用层攻击" ;;
        4.29) page=317; title="检查是否具有用户访问权限统一管理功能" ;;
        4.30) page=318; title="检查是否能够设置统一管理措施，是否对远程管理进行限制" ;;
        4.31) page=321; title="检查是否能够设置最大并发会话连接数、会话建立速率、单用户并发会话数" ;;
        4.32) page=324; title="检查所有应用是否具备备份与恢复功能（指导书4.33为\"是否基于国产自主可控软硬件自主开发\"，与本项编号存在1位偏差，请人工核对）" ;;
        5.*) page=327; title="第5章 网络安全（指导书本章为空白章节，未列出具体条款编号，以下网络设备类核查项供参考，均标记不适用）" ;;
        6.1) page=327; title="使用的网络设备、服务器、终端等，应选用进入《全军计算机及网络设备集中采购目录》的产品" ;;
        6.2) page=329; title="使用的安全网关、防火墙等信息安全产品，应通过信息安全测评认证机构的认证" ;;
        6.3) page=330; title="机房应有序、规范、合理走线布线，明确互联网区域和内部区域，粘贴标识，区分不同线路" ;;
        6.4) page=331; title="机房应符合GB 2887-2011中4.6.1所规定的温湿度等要求" ;;
        6.5) page=332; title="机房应安装安防监控设备，对机房的人员进出、设备操作等进行监管" ;;
        6.6) page=333; title="机房应将涉密区域和互联网区域设置于不同场所" ;;
        6.7) page=333; title="在涉密网络中使用过的打印机、复印机、刻录机、扫描仪、存储载体等设备严禁在互联网中使用" ;;
        6.8) page=334; title="应选用《关键软硬件自主可控产品名录》中的芯片类、计算机及外设类、网络设备类、安全防护设备类、存储设备类等产品" ;;
        6.9) page=335; title="与安全防护等级低的网络使用不同色系线缆进行严格隔离区分" ;;
        7.1) page=336; title="应设置安全管理机构，保证系统安全措施的落实" ;;
        7.2) page=337; title="应配备专职系统安全保密管理人员，负责系统安全措施的落实" ;;
        7.3) page=337; title="应具有安全管理领导机构，督导系统安全措施的落实" ;;
        7.4) page=338; title="应具有系统安全保密技术管理人员，具体负责安全保密技术措施的落实" ;;
        7.5) page=339; title="应具有完善的应急响应体系，应对突发事件" ;;
        8.1) page=340; title="应具有日常安全管理制度、入网审批制度和系统安全保密检查制度等" ;;
        8.2) page=341; title="应具有安全管理操作规程、安全监控操作规程、安全审计操作规程、应急响应操作规程等" ;;
        8.3) page=341; title="应具有日常系统备份制度和存储载体使用管理制度" ;;
        8.4) page=342; title="应具有脆弱性分析等规程" ;;
        9.1) page=343; title="应具有应急响应预案，发生危及系统安全的事件时应根据操作规程及时采取措施" ;;
        9.2) page=344; title="安全保密管理人员应能对网络安全防护设备进行统一配置、管理" ;;
        9.3) page=345; title="应具有与实际情况相符且完整的安全保密策略文档和安全保密技术及产品配置的详细记录" ;;
        9.4) page=346; title="安全保密技术管理人员应对每日网络和系统运行情况实施安全审计，每月组织安全性检测" ;;
        10.1|10.1_m) page=347; title="第10章 协议安全审计：审计协议的机密性、完整性、认可性、不可否认性（指导书本章未细分编号）" ;;
        *) page=""; title="" ;;
    esac
    if [ -n "$page" ]; then
        echo "《配置核查作业指导书》第${page}页：${title}"
    else
        echo ""
    fi
}

add_result() {
    # $1 id  $2 category  $3 title  $4 status(pass/fail/manual/na)  $5 detail  $6 chapter  $7 recommendation
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"
    R_CAT[$R_COUNT]="$2"
    R_TITLE[$R_COUNT]="$3"
    R_STATUS[$R_COUNT]="$4"
    R_DETAIL[$R_COUNT]="$5"
    R_CHAPTER[$R_COUNT]="$6"
    R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
}

run_cmd() {
    timeout 6 bash -c "$1" 2>/dev/null
}

os_release_field() {
    awk -F= -v f="$1" '$1==f{gsub(/"/,"",$2); print $2}' /etc/os-release 2>/dev/null
}

pkg_installed() {
    if [ "$PKG_MGR" = "apt" ]; then
        dpkg -s "$1" >/dev/null 2>&1
    else
        rpm -q "$1" >/dev/null 2>&1
    fi
}

svc_active() {
    systemctl is-active "$1" >/dev/null 2>&1
}

svc_enabled() {
    systemctl is-enabled "$1" >/dev/null 2>&1
}

# ---------- 发行版/包管理器检测 ----------
PKG_MGR=""
if command -v apt-get >/dev/null 2>&1; then
    PKG_MGR="apt"
elif command -v yum >/dev/null 2>&1 || command -v dnf >/dev/null 2>&1; then
    PKG_MGR="yum"
fi

OS_NAME="$(os_release_field NAME)"
OS_VER="$(os_release_field VERSION)"
OS_ID="$(os_release_field ID)"
KERNEL_VER="$(uname -r)"
HOSTNAME_STR="$(hostname 2>/dev/null)"

KYLIN_TYPE="未知/其他Linux"
case "$OS_ID" in
    kylin) KYLIN_TYPE="银河麒麟 Kylin OS（apt/dpkg）" ;;
    neokylin) KYLIN_TYPE="中标麒麟 NeoKylin（yum/rpm）" ;;
    *)
        if [ -f /etc/kylin-release ]; then
            KYLIN_TYPE="银河麒麟 Kylin OS（apt/dpkg，来自/etc/kylin-release）"
        elif [ -f /etc/neokylin-release ]; then
            KYLIN_TYPE="中标麒麟 NeoKylin（yum/rpm，来自/etc/neokylin-release）"
        elif [ "$PKG_MGR" = "apt" ]; then
            KYLIN_TYPE="疑似银河麒麟/Debian系（apt/dpkg）"
        elif [ "$PKG_MGR" = "yum" ]; then
            KYLIN_TYPE="疑似中标麒麟/RHEL系（yum/rpm）"
        fi
        ;;
esac

IS_ROOT=0
[ "$(id -u)" = "0" ] && IS_ROOT=1

echo "================================================================"
echo "  配置核查工具 - 中标麒麟/银河麒麟版（无需Python）"
echo "  参考标准：配置核查作业指导书正式版2026_4_1"
echo "================================================================"
echo "  系统：$OS_NAME $OS_VER  ($KYLIN_TYPE)"
echo "  内核：$KERNEL_VER"
if [ "$IS_ROOT" -ne 1 ]; then
    echo "  [警告] 当前非root运行，部分检查结果可能不准确，建议 sudo 执行。"
fi
echo ""
echo "开始检查，请稍候..."
echo ""

# ============================================================
# 第1章 系统安全
# ============================================================

check_1_1_patch() {
    local detail=""
    if [ "$PKG_MGR" = "apt" ]; then
        local cnt
        cnt="$(run_cmd "apt list --upgradable 2>/dev/null | grep -v '^Listing' | wc -l")"
        cnt="${cnt:-0}"
        if [ "$cnt" -gt 0 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "fail" "检测到 $cnt 个可升级的软件包（apt list --upgradable），系统补丁未安装到最新。" "第1章" "使用 apt update && apt upgrade 及时安装安全补丁。"
        else
            add_result "1.1" "系统安全" "补丁安装情况" "pass" "apt list --upgradable 未发现可升级的软件包，或无法连接更新源（离线环境常见）。" "第1章" "定期检查并安装系统安全补丁。"
        fi
    else
        local out
        out="$(run_cmd "yum check-update 2>/dev/null | grep -Ev '^$|^Loaded|^Last metadata|^Obsoleting|^Security:' | wc -l")"
        out="${out:-0}"
        if [ "$out" -gt 0 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "fail" "yum check-update 检测到约 $out 个可更新软件包，系统补丁未安装到最新。" "第1章" "使用 yum update 及时安装安全补丁。"
        else
            add_result "1.1" "系统安全" "补丁安装情况" "pass" "yum check-update 未发现可更新软件包，或无法连接更新源（离线环境常见）。" "第1章" "定期检查并安装系统安全补丁。"
        fi
    fi
}

check_1_2_antivirus() {
    local found=""
    for p in clamav clamd 360wangzhan huorong-eds sangfor-eds; do
        if pkg_installed "$p" 2>/dev/null || svc_active "$p" 2>/dev/null; then
            found="$found $p"
        fi
    done
    if [ -n "$found" ]; then
        add_result "1.2" "系统安全" "防病毒软件" "pass" "检测到已安装/运行的安全防护软件：$found" "第1章" "保持防病毒/终端安全软件病毒库及策略为最新。"
    else
        add_result "1.2" "系统安全" "防病毒软件" "manual" "未检测到常见国产终端安全软件特征，请人工确认是否安装了防病毒/EDR软件（如天融信、奇安信、深信服终端等）。" "第1章" "根据单位要求安装终端安全防护软件。"
    fi
}

check_1_3_services() {
    local risky="telnet rsh rlogin tftp vsftpd rexec"
    local found=""
    for s in $risky; do
        if svc_active "$s" 2>/dev/null; then
            found="$found $s"
        fi
    done
    if [ -n "$found" ]; then
        add_result "1.3" "系统安全" "高危服务/端口" "fail" "检测到以下高危服务正在运行：$found" "第1章" "停用 telnet/rsh/rlogin/tftp 等明文/高危服务，使用 SSH 替代。"
    else
        add_result "1.3" "系统安全" "高危服务/端口" "pass" "未检测到 telnet/rsh/rlogin/tftp/rexec 等高危服务处于运行状态。" "第1章" "定期核查系统服务，禁用不必要的高危服务。"
    fi
}

check_1_4_firewall() {
    if svc_active firewalld; then
        local zone
        zone="$(run_cmd 'firewall-cmd --get-default-zone')"
        add_result "1.4" "系统安全" "防火墙状态" "pass" "firewalld 正在运行，默认区域：${zone:-未知}" "第1章" "保持防火墙开启，按最小化原则配置放行规则。"
    elif svc_active ufw; then
        local st
        st="$(run_cmd 'ufw status | head -1')"
        add_result "1.4" "系统安全" "防火墙状态" "pass" "ufw 正在运行：${st:-active}" "第1章" "保持防火墙开启，按最小化原则配置放行规则。"
    elif command -v iptables >/dev/null 2>&1; then
        local rules
        rules="$(run_cmd 'iptables -S | wc -l')"
        if [ "${rules:-0}" -gt 3 ] 2>/dev/null; then
            add_result "1.4" "系统安全" "防火墙状态" "pass" "firewalld/ufw 未运行，但检测到 iptables 已配置 ${rules} 条规则。" "第1章" "确保 iptables 规则集在重启后持久化生效。"
        else
            add_result "1.4" "系统安全" "防火墙状态" "fail" "未检测到 firewalld/ufw 运行，且 iptables 规则为空或默认放行。" "第1章" "启用 firewalld/ufw 或配置 iptables 规则，禁止非必要访问。"
        fi
    else
        add_result "1.4" "系统安全" "防火墙状态" "manual" "未检测到 firewalld/ufw/iptables，请人工核实主机防火墙配置。" "第1章" "配置主机防火墙，限制非必要的入站/出站访问。"
    fi
}

check_1_5_selinux() {
    if command -v getenforce >/dev/null 2>&1; then
        local st
        st="$(run_cmd getenforce)"
        if [ "$st" = "Enforcing" ]; then
            add_result "1.5" "系统安全" "SELinux/强制访问控制" "pass" "SELinux 状态为 Enforcing。" "第1章" "保持 SELinux 处于 Enforcing 模式，不建议关闭。"
        elif [ "$st" = "Permissive" ]; then
            add_result "1.5" "系统安全" "SELinux/强制访问控制" "fail" "SELinux 状态为 Permissive（仅记录不拦截）。" "第1章" "将 SELinux 设置为 Enforcing 模式。"
        else
            add_result "1.5" "系统安全" "SELinux/强制访问控制" "fail" "SELinux 状态为 Disabled。" "第1章" "启用 SELinux 并设置为 Enforcing 模式。"
        fi
    elif command -v aa-status >/dev/null 2>&1; then
        local out
        out="$(run_cmd 'aa-status --enabled && echo enabled || echo disabled')"
        if echo "$out" | grep -q enabled; then
            add_result "1.5" "系统安全" "SELinux/强制访问控制" "pass" "AppArmor 已启用（银河麒麟常用强制访问控制机制）。" "第1章" "保持 AppArmor 启用并维护应用策略。"
        else
            add_result "1.5" "系统安全" "SELinux/强制访问控制" "fail" "AppArmor 未启用。" "第1章" "启用 AppArmor 强制访问控制。"
        fi
    else
        add_result "1.5" "系统安全" "SELinux/强制访问控制" "manual" "未检测到 SELinux/AppArmor 工具，请人工核实强制访问控制机制状态。" "第1章" "启用系统自带的强制访问控制机制（SELinux/AppArmor）。"
    fi
}

check_1_6_ssh_crypto() {
    local cfg="/etc/ssh/sshd_config"
    if [ -f "$cfg" ]; then
        local proto weak
        weak="$(grep -Ei '^\s*Ciphers' "$cfg" | grep -Ei '3des|arcfour|rc4|blowfish|cbc')"
        if [ -n "$weak" ]; then
            add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "fail" "sshd_config 中配置了弱加密算法：$weak" "第1章" "禁用 3DES/RC4/CBC 等弱加密套件，使用 AES-GCM/ChaCha20 等强算法。"
        else
            add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "pass" "sshd_config 未显式配置已知弱加密算法（Ciphers未设置时使用系统默认强套件）。" "第1章" "定期核查 SSH 加密套件配置，避免使用弱算法。"
        fi
    else
        add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "na" "未检测到 sshd 配置文件，可能未安装 SSH 服务。" "第1章" "如需远程管理，安装并正确配置 OpenSSH。"
    fi
}

check_1_17_redis() {
    if svc_active redis || svc_active redis-server || pkg_installed redis 2>/dev/null; then
        local conf=""
        for c in /etc/redis/redis.conf /etc/redis.conf; do
            [ -f "$c" ] && conf="$c" && break
        done
        if [ -n "$conf" ]; then
            local bind pass
            bind="$(grep -Ei '^\s*bind' "$conf" | head -1)"
            pass="$(grep -Ei '^\s*requirepass' "$conf" | head -1)"
            if [ -z "$pass" ]; then
                add_result "1.17" "系统安全" "Redis安全配置" "fail" "Redis 配置文件 $conf 未设置 requirepass（无密码认证）。绑定：${bind:-未设置}" "第1章" "为 Redis 配置 requirepass 强密码，并将 bind 限制为内网地址。"
            else
                add_result "1.17" "系统安全" "Redis安全配置" "pass" "Redis 已配置 requirepass 密码认证。绑定：${bind:-未设置}" "第1章" "定期更换 Redis 密码，避免绑定 0.0.0.0 对外暴露。"
            fi
        else
            add_result "1.17" "系统安全" "Redis安全配置" "manual" "检测到 Redis 服务运行，但未找到标准配置文件路径，请人工核查。" "第1章" "核查 Redis 认证与绑定地址配置。"
        fi
    else
        add_result "1.17" "系统安全" "Redis安全配置" "na" "未检测到 Redis 服务运行，标记不适用。" "第1章" "如部署 Redis，需配置密码认证并限制访问来源。"
    fi
}

check_1_24_auditpolicy() {
    if svc_active auditd; then
        local rules
        rules="$(run_cmd 'auditctl -l 2>/dev/null | wc -l')"
        if [ "${rules:-0}" -gt 0 ] 2>/dev/null; then
            add_result "1.24" "系统安全" "审计策略" "pass" "auditd 正在运行，已加载 $rules 条审计规则。" "第1章" "根据合规要求持续完善审计规则覆盖面。"
        else
            add_result "1.24" "系统安全" "审计策略" "fail" "auditd 正在运行，但未加载任何审计规则。" "第1章" "配置 /etc/audit/rules.d 下的审计规则并重载 auditd。"
        fi
    elif svc_active systemd-journald; then
        add_result "1.24" "系统安全" "审计策略" "manual" "未运行 auditd，仅有 systemd-journald 日志，请人工确认审计策略是否满足要求。" "第1章" "安装并启用 auditd，配置符合合规要求的审计规则。"
    else
        add_result "1.24" "系统安全" "审计策略" "fail" "未检测到 auditd 或有效的审计日志机制。" "第1章" "安装并启用 auditd 审计服务。"
    fi
}

check_1_25_defender() {
    add_result "1.25" "系统安全" "系统自带安全防护组件" "na" "Windows Defender 为Windows专有组件，Linux平台不适用；请参照1.2防病毒/EDR核查结果。" "第1章" "参考本报告1.2项，确认已部署等效的终端安全防护能力。"
}

check_1_18_resourcelimits() {
    local nofile
    nofile="$(run_cmd 'ulimit -n')"
    local limitsconf="/etc/security/limits.conf"
    if [ -f "$limitsconf" ] && grep -Eq '^\s*\*.*nofile' "$limitsconf" 2>/dev/null; then
        add_result "1.18" "系统安全" "系统资源限制" "pass" "当前 nofile 软限制为 ${nofile:-未知}，/etc/security/limits.conf 已配置资源限制。" "第1章" "根据业务负载合理设置进程/用户资源限制，防止资源耗尽型攻击。"
    else
        add_result "1.18" "系统安全" "系统资源限制" "manual" "当前 nofile 软限制为 ${nofile:-未知}，未在 limits.conf 中发现全局限制配置，请人工核实是否满足需求。" "第1章" "在 /etc/security/limits.conf 中配置合理的资源限制。"
    fi
}

check_1_26_securitylogs() {
    local logdir="/var/log"
    local total
    total="$(run_cmd "du -sh $logdir 2>/dev/null | awk '{print \$1}'")"
    if svc_active rsyslog || svc_active syslog-ng || [ -f /etc/logrotate.conf ]; then
        add_result "1.26" "系统安全" "安全日志容量与留存" "pass" "日志服务（rsyslog/syslog-ng）与 logrotate 轮转机制存在，/var/log 当前占用 ${total:-未知}。" "第1章" "确保日志留存时长满足合规要求（通常不少于6个月）。"
    else
        add_result "1.26" "系统安全" "安全日志容量与留存" "manual" "未检测到标准日志轮转机制，/var/log 当前占用 ${total:-未知}，请人工确认日志留存策略。" "第1章" "配置 logrotate 或等效机制，确保日志不因容量不足被覆盖。"
    fi
}

# ---------- 数据库相关（1.7 - 1.23）----------
DB_TYPE=""
detect_db() {
    if svc_active mysqld || svc_active mysql || svc_active mariadb; then
        DB_TYPE="MySQL/MariaDB"
    elif svc_active DmServiceDMSERVER || pgrep -f dmserver >/dev/null 2>&1; then
        DB_TYPE="达梦数据库(DM)"
    elif pgrep -f kingbase >/dev/null 2>&1 || svc_active kingbase; then
        DB_TYPE="人大金仓(Kingbase)"
    elif svc_active postgresql; then
        DB_TYPE="PostgreSQL"
    else
        DB_TYPE=""
    fi
}

db_na_or_manual() {
    # $1 id $2 title $3 rec
    if [ -z "$DB_TYPE" ]; then
        add_result "$1" "系统安全-数据库" "$2" "na" "未检测到本机运行数据库服务（MySQL/达梦/Kingbase/PostgreSQL），标记不适用。" "第1章" "$3"
    else
        add_result "$1" "系统安全-数据库" "$2" "manual" "检测到数据库服务：$DB_TYPE。该项需使用数据库管理员账户登录后核查，脚本未内置数据库凭据，请人工核查。" "第1章" "$3"
    fi
}

check_1_7_dbaccounts()   { db_na_or_manual "1.7"  "数据库账户管理"       "核查数据库账户是否唯一分配、无共享账户，删除或锁定无用账户。"; }
check_1_8_dbprocedures() { db_na_or_manual "1.8"  "数据库存储过程管理"   "核查高危存储过程（如xp_cmdshell等效功能）是否禁用。"; }
check_1_9_dbgrants()     { db_na_or_manual "1.9"  "数据库权限最小化"     "按最小权限原则分配数据库账户权限，避免使用超级管理员账户运行业务。"; }

check_1_10_dbaccess() {
    detect_db
    if [ -z "$DB_TYPE" ]; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "na" "未检测到本机运行数据库服务，标记不适用。" "第1章" "限制数据库监听地址与访问来源IP。"
        return
    fi
    local bind=""
    for c in /etc/mysql/my.cnf /etc/my.cnf /etc/mysql/mysql.conf.d/mysqld.cnf; do
        [ -f "$c" ] && bind="$bind $(grep -Ei '^\s*bind-address' "$c" 2>/dev/null)"
    done
    local listen
    listen="$(run_cmd "ss -lntp 2>/dev/null | grep -E ':3306|:5236|:54321'")"
    if echo "$listen" | grep -q '0\.0\.0\.0\|\*:'; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "fail" "数据库端口监听在 0.0.0.0（对外暴露）：$listen" "第1章" "将数据库绑定地址改为内网/127.0.0.1，并通过防火墙限制访问来源IP。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "数据库监听地址未发现对公网暴露的0.0.0.0绑定。绑定配置：${bind:-默认}" "第1章" "持续保持数据库仅对可信来源开放访问。"
    fi
}

check_1_11_dbbackup() {
    detect_db
    if [ -z "$DB_TYPE" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "na" "未检测到本机运行数据库服务，标记不适用。" "第1章" "建立数据库定期备份机制并定期验证可恢复性。"
        return
    fi
    local cron
    cron="$(run_cmd "grep -ril 'mysqldump\|dmap\|sys_dump\|backup' /etc/cron.d /etc/cron.daily /var/spool/cron 2>/dev/null")"
    if [ -n "$cron" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "检测到与数据库备份相关的计划任务：$cron" "第1章" "定期验证备份文件的可恢复性，并异地存储备份。"
    else
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "manual" "检测到数据库服务($DB_TYPE)，但未在常见crontab位置发现自动备份任务，请人工核实备份策略。" "第1章" "建立数据库定期备份机制（如mysqldump/达梦dmap/金仓sys_dump）。"
    fi
}

check_1_12_dbauditplugin() {
    detect_db
    if [ "$DB_TYPE" = "MySQL/MariaDB" ]; then
        local glog=""
        for c in /etc/mysql/my.cnf /etc/my.cnf /etc/mysql/mysql.conf.d/mysqld.cnf; do
            [ -f "$c" ] && glog="$glog $(grep -Ei '^\s*(general_log|audit_log|server_audit)' "$c" 2>/dev/null)"
        done
        if [ -n "$(echo "$glog" | xargs)" ]; then
            add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "MySQL 配置文件中检测到审计/日志相关配置：$glog" "第1章" "确保审计日志覆盖登录、权限变更、敏感数据操作等关键行为。"
        else
            add_result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "MySQL 配置文件中未发现 general_log/audit_log 等审计相关配置。" "第1章" "启用 MySQL 审计插件（如audit_log/server_audit）记录关键操作。"
        fi
    else
        db_na_or_manual "1.12" "数据库审计插件" "启用数据库自带审计功能（达梦审计策略/金仓审计）记录关键操作。"
    fi
}

check_1_13_dbisolation() { db_na_or_manual "1.13" "数据库与业务隔离" "评估数据库是否与Web/应用服务同机部署，建议物理或网络隔离。"; }

check_1_14_middleware() {
    local found=""
    for s in nginx httpd apache2 tomcat weblogic; do
        if svc_active "$s" 2>/dev/null || pkg_installed "$s" 2>/dev/null; then
            found="$found $s"
        fi
    done
    if pgrep -f tomcat >/dev/null 2>&1; then found="$found tomcat(进程)"; fi
    if [ -n "$found" ]; then
        add_result "1.14" "系统安全" "中间件安全配置" "manual" "检测到中间件：$found，请人工核查版本、默认页面、目录浏览、错误信息回显等配置。" "第1章" "关闭默认页面/目录浏览/详细错误信息，及时更新中间件版本。"
    else
        add_result "1.14" "系统安全" "中间件安全配置" "na" "未检测到常见中间件（nginx/httpd/tomcat/weblogic）运行，标记不适用。" "第1章" "如部署中间件，需加固默认配置并及时更新版本。"
    fi
}

check_1_15_dbremote() {
    detect_db
    if [ -z "$DB_TYPE" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "na" "未检测到本机运行数据库服务，标记不适用。" "第1章" "限制数据库仅允许可信运维终端远程访问。"
        return
    fi
    local rule
    rule="$(run_cmd "iptables -S 2>/dev/null | grep -E '3306|5236|54321'")"
    if [ -n "$rule" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "检测到针对数据库端口的防火墙规则：$rule" "第1章" "持续核查允许访问数据库的来源地址清单，及时清理失效规则。"
    else
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "检测到数据库服务($DB_TYPE)，但未在iptables中发现针对性访问控制规则，请人工核实firewalld/安全组配置。" "第1章" "通过防火墙/安全组限制数据库远程访问来源。"
    fi
}

check_1_16_17_inputcheck() { add_result "1.16-17" "系统安全" "输入验证/SQL注入防护" "manual" "输入验证与SQL注入防护属应用层代码逻辑，需人工/工具（如AWVS）核查应用系统。" "第1章" "应用层对用户输入进行合法性校验，使用参数化查询防止SQL注入。"; }

check_1_19_dbdefaults() {
    detect_db
    if [ -z "$DB_TYPE" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "na" "未检测到本机运行数据库服务，标记不适用。" "第1章" "修改数据库默认端口，禁用/改名默认账户。"
        return
    fi
    local port_default
    port_default="$(run_cmd "ss -lntp 2>/dev/null | grep -E ':3306|:5236|:54321'")"
    if [ -n "$port_default" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "数据库仍使用默认端口对外监听：$port_default" "第1章" "修改数据库默认监听端口，并核查root/sys等默认账户是否已加固。"
    else
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "manual" "未在默认端口监听，请人工核实数据库默认账户是否已重命名/禁用/改密。" "第1章" "禁用或加固数据库默认账户，修改默认端口。"
    fi
}

check_1_20_dbpolicy() { db_na_or_manual "1.20" "数据库安全策略" "制定并落实数据库密码策略、访问策略等安全配置基线。"; }
check_1_21_dbaudit()  { db_na_or_manual "1.21" "数据库操作审计" "开启数据库操作审计功能，记录敏感操作行为。"; }
check_1_22_dbauditretention() { db_na_or_manual "1.22" "数据库审计日志留存" "确保数据库审计日志留存时长满足合规要求（通常不少于6个月）。"; }
check_1_23_dbappsep()  { db_na_or_manual "1.23" "应用与数据库账户分离" "应用系统连接数据库应使用独立受限账户，避免使用管理员账户。"; }

# ============================================================
# 第2章 用户安全
# ============================================================

check_2_1_passwordexpiry() {
    local maxdays
    maxdays="$(grep -E '^\s*PASS_MAX_DAYS' /etc/login.defs 2>/dev/null | awk '{print $2}')"
    if [ -n "$maxdays" ] && [ "$maxdays" -gt 0 ] 2>/dev/null && [ "$maxdays" -le 90 ] 2>/dev/null; then
        add_result "2.1" "用户安全" "密码有效期" "pass" "/etc/login.defs 中 PASS_MAX_DAYS=$maxdays（<=90天）。" "第2章" "定期强制用户更换密码，建议不超过90天。"
    else
        add_result "2.1" "用户安全" "密码有效期" "fail" "/etc/login.defs 中 PASS_MAX_DAYS=${maxdays:-未设置}，未配置或超过90天。" "第2章" "在 /etc/login.defs 设置 PASS_MAX_DAYS 不超过90天。"
    fi
}

check_2_4_uniqueaccounts() {
    if [ ! -r /etc/passwd ]; then
        add_result "2.4" "用户安全" "账户唯一性" "manual" "无法读取 /etc/passwd，请人工核实账户唯一性。" "第2章" "确保每个账户UID唯一，避免共享账户。"
        return
    fi
    local dup
    dup="$(awk -F: '{print $3}' /etc/passwd 2>/dev/null | sort | uniq -d)"
    if [ -n "$dup" ]; then
        add_result "2.4" "用户安全" "账户唯一性" "fail" "/etc/passwd 中发现重复UID：$dup" "第2章" "确保每个账户UID唯一，避免共享账户。"
    else
        add_result "2.4" "用户安全" "账户唯一性" "pass" "/etc/passwd 中未发现重复UID，账户具备唯一性。" "第2章" "持续避免创建共享账户，人员离职及时删除/锁定账户。"
    fi
}

check_2_5_redundantservices() {
    local enabled
    enabled="$(run_cmd 'systemctl list-unit-files --state=enabled --type=service 2>/dev/null | wc -l')"
    add_result "2.5" "用户安全" "多余服务清理" "manual" "当前系统开机自启服务数：${enabled:-未知}，请人工核实是否存在未使用的多余服务/账户。" "第2章" "定期核查并清理不再使用的服务、账户与软件。"
}

check_2_2_userpatch() {
    add_result "2.2" "用户安全" "终端补丁管理" "manual" "该项与1.1系统补丁核查内容重叠，请以1.1核查结果为准，并结合终端管理平台核实补丁分发情况。" "第2章" "通过补丁管理平台统一分发、核查终端补丁安装情况。"
}

check_2_6_illegalinternet() {
    local est
    est="$(run_cmd "ss -ntp state established 2>/dev/null | wc -l")"
    add_result "2.6" "用户安全" "违规外联检测" "manual" "当前已建立的网络连接数：${est:-未知}，是否存在违规外联需结合出网流量/日志人工核查。" "第2章" "通过防火墙出站策略与流量审计，发现并阻断违规外联行为。"
}

check_2_8_endpointcontrol() {
    add_result "2.8" "用户安全" "终端管控软件" "manual" "该项需人工确认是否安装了统一终端安全管理（EMM/终端管控）软件。" "第2章" "部署统一终端安全管理软件，落实准入、补丁、外联等管控策略。"
}

check_2_15_securitypolicy() {
    add_result "2.15" "用户安全" "安全策略落实" "manual" "安全策略（如屏幕锁定、剪切板管控等）需结合终端管理平台/桌面环境配置人工核查。" "第2章" "落实屏幕锁定超时、剪切板等桌面安全策略。"
}

check_2_16_patchlatest() {
    local kern mtime
    kern="$(uname -r)"
    if [ "$PKG_MGR" = "apt" ]; then
        mtime="$(stat -c %Y /var/lib/apt/periodic/update-success-stamp 2>/dev/null)"
        [ -z "$mtime" ] && mtime="$(stat -c %Y /var/log/apt/history.log 2>/dev/null)"
    else
        mtime="$(stat -c %Y /var/lib/rpm/Packages 2>/dev/null)"
        [ -z "$mtime" ] && mtime="$(stat -c %Y /var/log/yum.log 2>/dev/null)"
    fi
    if [ -n "$mtime" ]; then
        local now days
        now="$(date +%s)"
        days=$(( (now - mtime) / 86400 ))
        if [ "$days" -gt 90 ]; then
            add_result "2.16" "用户安全" "补丁是否为最新版本" "fail" "当前内核版本：$kern。距上次系统更新已 $days 天（超过90天/约3个月）。" "第2章" "定期执行系统更新，确保补丁不超过3个月未更新。"
        else
            add_result "2.16" "用户安全" "补丁是否为最新版本" "pass" "当前内核版本：$kern。距上次系统更新 $days 天，未超过3个月。" "第2章" "跟踪厂商安全公告，及时升级到推荐的最新补丁版本。"
        fi
    else
        add_result "2.16" "用户安全" "补丁是否为最新版本" "manual" "当前内核版本：$kern。无法获取上次更新时间戳，是否为厂商最新推荐版本需结合官方安全公告人工核实。" "第2章" "跟踪厂商安全公告，及时升级到推荐的最新补丁版本。"
    fi
}

check_2_7_fstype() {
    local fstype
    fstype="$(run_cmd "df -T / 2>/dev/null | tail -1 | awk '{print \$2}'")"
    add_result "2.7" "用户安全" "文件系统安全" "na" "Windows NTFS检查在Linux平台不适用；当前根分区文件系统为：${fstype:-未知}（ext4/xfs原生支持权限位与ACL）。" "第2章" "确保关键分区文件系统支持权限位/ACL，并正确设置目录权限。"
}

check_2_9_riskysoftware() {
    local found=""
    for p in teamviewer anydesk todesk sunlogin; do
        if pkg_installed "$p" 2>/dev/null || pgrep -f "$p" >/dev/null 2>&1; then
            found="$found $p"
        fi
    done
    if [ -n "$found" ]; then
        add_result "2.9" "用户安全" "高风险软件" "fail" "检测到高风险远程控制软件：$found" "第2章" "卸载TeamViewer/AnyDesk/向日葵等未经授权的远程控制软件。"
    else
        add_result "2.9" "用户安全" "高风险软件" "pass" "未检测到常见高风险远程控制软件。" "第2章" "定期核查并卸载未经授权的高风险软件。"
    fi
}

check_2_10_usb() {
    if pkg_installed usbguard 2>/dev/null && svc_active usbguard; then
        add_result "2.10" "用户安全" "USB外设管控" "pass" "检测到 usbguard 服务正在运行，具备USB设备管控能力。" "第2章" "持续维护usbguard白名单策略。"
    else
        local udevrule
        udevrule="$(run_cmd "grep -ril 'usb_storage\|authorized' /etc/udev/rules.d 2>/dev/null")"
        if [ -n "$udevrule" ]; then
            add_result "2.10" "用户安全" "USB外设管控" "pass" "检测到udev规则对USB存储设备进行管控：$udevrule" "第2章" "定期核查udev/usbguard规则的有效性。"
        else
            add_result "2.10" "用户安全" "USB外设管控" "fail" "未检测到usbguard或udev层面的USB存储管控规则。" "第2章" "部署usbguard或配置udev规则禁用/管控USB存储设备。"
        fi
    fi
}

check_2_11_passwordpolicy() {
    local pamfile=""
    for f in /etc/pam.d/system-auth /etc/pam.d/common-password; do
        [ -f "$f" ] && pamfile="$f" && break
    done
    if [ -n "$pamfile" ]; then
        local pwquality minlen
        minlen="$(grep -E 'pam_pwquality|pam_cracklib' "$pamfile" | grep -oE 'minlen=[0-9]+' | head -1)"
        if [ -n "$minlen" ]; then
            add_result "2.11" "用户安全" "口令复杂度策略" "pass" "$pamfile 中配置了密码复杂度策略：$minlen" "第2章" "确保密码长度不少于8位，包含大小写字母、数字、特殊字符中至少3类。"
        else
            add_result "2.11" "用户安全" "口令复杂度策略" "fail" "$pamfile 中未发现 pam_pwquality/pam_cracklib 的minlen等复杂度配置。" "第2章" "在PAM中启用pam_pwquality，配置密码长度与复杂度要求。"
        fi
    else
        add_result "2.11" "用户安全" "口令复杂度策略" "manual" "未找到标准PAM密码策略配置文件，请人工核实密码复杂度要求。" "第2章" "配置PAM密码复杂度策略（长度、字符种类、历史）。"
    fi
}

check_2_12_wireless() {
    if command -v nmcli >/dev/null 2>&1; then
        local wifi
        wifi="$(run_cmd 'nmcli radio wifi')"
        if [ "$wifi" = "enabled" ]; then
            add_result "2.12" "用户安全" "无线网络管控" "manual" "检测到WiFi无线功能已启用（nmcli radio wifi=enabled），请人工核实是否符合无线管控要求。" "第2章" "服务器/涉密终端建议关闭无线网卡功能。"
        else
            add_result "2.12" "用户安全" "无线网络管控" "pass" "WiFi无线功能已关闭（nmcli radio wifi=${wifi:-disabled}）。" "第2章" "持续保持不必要的无线功能关闭状态。"
        fi
    else
        add_result "2.12" "用户安全" "无线网络管控" "na" "未检测到 NetworkManager（nmcli），可能为无GUI服务器场景，标记不适用。" "第2章" "如具备无线网卡，需按管控要求限制使用。"
    fi
}

check_2_13_outboundcontrol() {
    local out
    out="$(run_cmd "iptables -S OUTPUT 2>/dev/null | wc -l")"
    if [ "${out:-0}" -gt 1 ] 2>/dev/null; then
        add_result "2.13" "用户安全" "外联控制" "pass" "检测到 iptables OUTPUT 链已配置 $out 条规则，具备外联控制能力。" "第2章" "持续核查出站规则，遵循最小化原则。"
    else
        add_result "2.13" "用户安全" "外联控制" "manual" "iptables OUTPUT 链未发现针对性规则，请人工核实是否通过其他手段（如firewalld rich rule、上网行为管理）控制外联。" "第2章" "通过防火墙/上网行为管理限制非必要的对外连接。"
    fi
}

check_2_14_useraudit() {
    if svc_active auditd || [ -f /var/log/secure ] || [ -f /var/log/auth.log ]; then
        add_result "2.14" "用户安全" "用户行为审计" "pass" "检测到系统级用户行为日志机制（auditd/secure/auth.log）。" "第2章" "确保用户登录、操作行为日志被完整记录并留存。"
    else
        add_result "2.14" "用户安全" "用户行为审计" "fail" "未检测到用户行为审计相关日志机制。" "第2章" "启用auditd并配置用户操作审计规则。"
    fi
}

# ============================================================
# 第3章 数据安全
# ============================================================

check_3_1_encryption() {
    if command -v cryptsetup >/dev/null 2>&1; then
        local luks
        luks="$(run_cmd "lsblk -o NAME,FSTYPE 2>/dev/null | grep -i crypto_LUKS")"
        if [ -n "$luks" ]; then
            add_result "3.1" "数据安全" "磁盘加密" "pass" "检测到LUKS加密分区：$luks" "第3章" "对存储敏感数据的分区持续保持磁盘加密。"
        else
            add_result "3.1" "数据安全" "磁盘加密" "manual" "已安装cryptsetup但未检测到LUKS加密分区，请人工确认是否需要磁盘加密。" "第3章" "对存储敏感数据的磁盘/分区启用LUKS加密。"
        fi
    else
        add_result "3.1" "数据安全" "磁盘加密" "manual" "未检测到cryptsetup，请人工确认磁盘加密（LUKS）部署情况。" "第3章" "对存储敏感数据的磁盘/分区启用LUKS加密。"
    fi
}

check_3_5_logprotection() {
    local perm
    perm="$(run_cmd "stat -c '%a' /var/log 2>/dev/null")"
    if [ -n "$perm" ] && [ "$perm" -le 750 ] 2>/dev/null; then
        add_result "3.5" "数据安全" "日志文件保护" "pass" "/var/log 权限为 $perm，非过度开放。" "第3章" "持续限制日志目录/文件权限，防止未授权访问或篡改。"
    else
        add_result "3.5" "数据安全" "日志文件保护" "fail" "/var/log 权限为 ${perm:-未知}，权限设置可能过于开放。" "第3章" "将日志目录权限收紧（如750），仅允许必要账户访问。"
    fi
}

check_3_10_shares() {
    local smb=""
    local nfs=""
    svc_active smb 2>/dev/null && smb="smb"
    svc_active smbd 2>/dev/null && smb="smbd"
    [ -f /etc/exports ] && [ -s /etc/exports ] && nfs="nfs"
    if [ -n "$smb" ] || [ -n "$nfs" ]; then
        add_result "3.10" "数据安全" "文件共享安全" "manual" "检测到文件共享服务：${smb} ${nfs}，请人工核查共享目录权限与访问控制列表。" "第3章" "共享目录应设置访问账户与权限控制，禁止匿名/来宾访问。"
    else
        add_result "3.10" "数据安全" "文件共享安全" "na" "未检测到Samba/NFS共享服务，标记不适用。" "第3章" "如需文件共享，应严格控制访问权限，禁止匿名访问。"
    fi
}

check_3_8_storagepartition() {
    local mounts
    mounts="$(run_cmd "df -h 2>/dev/null | grep -E '/var$|/home$|/data$' | wc -l")"
    if [ "${mounts:-0}" -gt 0 ] 2>/dev/null; then
        add_result "3.8" "数据安全" "存储分区规划" "pass" "检测到 /var /home /data 等业务/日志目录使用独立分区，共 $mounts 个。" "第3章" "保持关键目录独立分区，避免单点磁盘写满影响系统运行。"
    else
        add_result "3.8" "数据安全" "存储分区规划" "manual" "未检测到 /var /home /data 独立分区，请人工确认存储分区规划是否合理。" "第3章" "建议将 /var、/home、业务数据目录规划为独立分区。"
    fi
}

# ============================================================
# 第4章 应用安全
# ============================================================

check_4_15_nla() {
    add_result "4.15" "应用安全" "网络级身份验证(NLA)" "na" "NLA为Windows RDP专有机制，Linux平台不适用。若部署了xrdp，请参照4.29远程管理项核查。" "第4章" "如部署图形化远程桌面（xrdp/VNC），应启用加密与身份验证。"
}

check_4_16_roleseparation() {
    local sudoers
    sudoers="$(run_cmd "wc -l < /etc/sudoers 2>/dev/null")"
    add_result "4.16" "应用安全" "三权分立" "manual" "sudo配置行数：${sudoers:-未知}，管理员/审计员/操作员的三权分立需人工核查账户与sudo授权划分。" "第4章" "划分系统管理、安全审计、业务操作等不同角色账户，避免权限过度集中。"
}

check_4_17_iprestrict() {
    local allow
    allow="$(run_cmd "cat /etc/hosts.allow 2>/dev/null | grep -v '^#' | grep -v '^$'")"
    if [ -n "$allow" ]; then
        add_result "4.17" "应用安全" "管理IP限制" "pass" "检测到 /etc/hosts.allow 配置了访问限制：$allow" "第4章" "持续维护允许管理访问的IP白名单。"
    else
        add_result "4.17" "应用安全" "管理IP限制" "manual" "/etc/hosts.allow 未配置限制规则，请人工核实是否通过firewalld/安全组等方式限制管理访问来源。" "第4章" "通过hosts.allow/firewalld/安全组限制管理访问来源IP。"
    fi
}

check_4_18_lockout() {
    local faillock=""
    for f in /etc/security/faillock.conf /etc/pam.d/system-auth /etc/pam.d/common-auth; do
        [ -f "$f" ] && grep -Eq 'pam_faillock|pam_tally2|deny=' "$f" 2>/dev/null && faillock="$f"
    done
    if [ -n "$faillock" ]; then
        add_result "4.18" "应用安全" "账户锁定策略" "pass" "$faillock 中检测到账户锁定策略配置（pam_faillock/pam_tally2）。" "第4章" "确保连续登录失败达到阈值后账户被锁定，锁定时间满足安全要求。"
    else
        add_result "4.18" "应用安全" "账户锁定策略" "fail" "未在PAM配置中发现pam_faillock/pam_tally2账户锁定策略。" "第4章" "在PAM中启用pam_faillock，配置登录失败锁定阈值。"
    fi
}

check_4_22_portseparation() { add_result "4.22" "应用安全" "管理端口与业务端口分离" "manual" "端口分离需结合具体业务架构人工核查（如管理端口是否与业务端口隔离监听）。" "第4章" "将管理端口与业务端口分离，管理端口仅限内网访问。"; }
check_4_23_appdbsep()      { add_result "4.23" "应用安全" "应用与数据库分离部署" "manual" "需人工核查应用服务与数据库是否部署于不同主机/网络区域。" "第4章" "应用系统与数据库建议分离部署，降低单点被攻破后的横向风险。"; }
check_4_24_appaudit()      { add_result "4.24" "应用安全" "应用操作审计" "manual" "应用层操作审计需结合具体应用系统日志人工核查。" "第4章" "应用系统应记录关键业务操作日志，满足可追溯要求。"; }

check_4_6_waf() {
    local found=""
    if pkg_installed modsecurity 2>/dev/null || pkg_installed libmodsecurity3 2>/dev/null; then
        found="modsecurity"
    fi
    if svc_active nginx 2>/dev/null && run_cmd "nginx -V 2>&1 | grep -qi modsecurity"; then
        found="$found nginx+modsecurity"
    fi
    if [ -n "$found" ]; then
        add_result "4.6" "应用安全" "Web应用防火墙(WAF)" "pass" "检测到WAF相关组件：$found" "第4章" "持续更新WAF规则库，覆盖OWASP Top10攻击类型。"
    else
        add_result "4.6" "应用安全" "Web应用防火墙(WAF)" "manual" "未检测到本机部署modsecurity等WAF组件，可能采用云端/前置WAF设备，请人工核实。" "第4章" "对外提供Web服务的系统应部署WAF防护。"
    fi
}

check_4_8_antitamper() {
    if pkg_installed aide 2>/dev/null; then
        add_result "4.8" "应用安全" "网页防篡改" "pass" "检测到 AIDE 文件完整性校验工具已安装。" "第4章" "定期更新AIDE基线库，及时发现文件被篡改的情况。"
    else
        add_result "4.8" "应用安全" "网页防篡改" "manual" "未检测到AIDE/tripwire等文件完整性校验工具，请人工核实是否部署了网页防篡改系统。" "第4章" "对外提供Web服务的系统应部署防篡改/文件完整性监控机制。"
    fi
}

check_4_12_sessionlimit() {
    local tmout
    tmout="$(grep -E '^\s*TMOUT=' /etc/profile /etc/profile.d/*.sh 2>/dev/null | head -1)"
    local sshalive
    sshalive="$(grep -Ei '^\s*ClientAliveInterval' /etc/ssh/sshd_config 2>/dev/null)"
    if [ -n "$tmout" ] || [ -n "$sshalive" ]; then
        add_result "4.12" "应用安全" "会话超时限制" "pass" "检测到会话超时配置：${tmout:-} ${sshalive:-}" "第4章" "保持登录会话空闲超时设置，建议不超过10分钟。"
    else
        add_result "4.12" "应用安全" "会话超时限制" "fail" "未检测到 TMOUT 或 sshd ClientAliveInterval 超时配置。" "第4章" "在/etc/profile设置TMOUT，并在sshd_config配置ClientAliveInterval。"
    fi
}

check_4_21_defaultport() {
    local port
    port="$(grep -Ei '^\s*Port' /etc/ssh/sshd_config 2>/dev/null | awk '{print $2}' | head -1)"
    port="${port:-22}"
    if [ "$port" = "22" ]; then
        add_result "4.21" "应用安全" "默认端口修改" "fail" "SSH仍使用默认端口22。" "第4章" "将SSH管理端口修改为非默认端口，降低被扫描/暴力破解风险。"
    else
        add_result "4.21" "应用安全" "默认端口修改" "pass" "SSH已修改为非默认端口：$port" "第4章" "持续保持管理端口非默认配置，并做好访问控制。"
    fi
}

check_4_29_remotemgmt() {
    local rootlogin
    rootlogin="$(grep -Ei '^\s*PermitRootLogin' /etc/ssh/sshd_config 2>/dev/null | awk '{print $2}' | head -1)"
    if [ "$rootlogin" = "no" ]; then
        add_result "4.29" "应用安全" "远程管理安全" "pass" "sshd_config 中 PermitRootLogin=no，已禁止root直接远程登录。" "第4章" "持续禁止root直接远程登录，使用普通账户+sudo方式管理。"
    else
        add_result "4.29" "应用安全" "远程管理安全" "fail" "sshd_config 中 PermitRootLogin=${rootlogin:-未设置（默认允许或prohibit-password）}，未明确禁止root远程登录。" "第4章" "在sshd_config设置 PermitRootLogin no，禁止root账户直接远程登录。"
    fi
}

check_4_31_appbackup() { add_result "4.31" "应用安全" "应用系统备份" "manual" "应用系统级备份策略需结合具体业务系统人工核查。" "第4章" "建立应用系统配置与数据的定期备份机制。"; }

check_4_32_domesticsoftware() {
    local dbkw="dmdbms kingbase gbase oceanbase tidb opengauss gaussdb"
    local mwkw="tongweb apusic primeton"
    local pkglist found=""
    if [ "$PKG_MGR" = "apt" ]; then
        pkglist="$(dpkg -l 2>/dev/null | awk '{print $2}')"
    else
        pkglist="$(rpm -qa 2>/dev/null)"
    fi
    local k
    for k in $dbkw $mwkw; do
        if echo "$pkglist" | grep -qi "$k"; then
            found="$found $k"
        fi
    done
    if [ -n "$found" ]; then
        add_result "4.32" "应用安全" "国产化软硬件适配" "pass" "操作系统本身即为国产操作系统：$KYLIN_TYPE；另检测到国产数据库/中间件相关软件包：$found。" "第4章" "持续核查关键应用软件与国产操作系统/CPU的适配兼容情况。"
    else
        add_result "4.32" "应用安全" "国产化软硬件适配" "pass" "操作系统本身即为国产操作系统：$KYLIN_TYPE。未检测到已安装的国产数据库/中间件软件包（如业务未部署此类组件，不影响判定）。" "第4章" "持续核查关键应用软件与国产操作系统/CPU的适配兼容情况。"
    fi
}

# ============================================================
# 第10章 密码与传输安全
# ============================================================

check_10_1_tls() {
    local weak
    weak="$(grep -Ei '^\s*Ciphers' /etc/ssh/sshd_config 2>/dev/null | grep -Ei 'ssl2|ssl3|tls1\.0|tls1\.1')"
    local opensslver
    opensslver="$(run_cmd 'openssl version')"
    if [ -n "$weak" ]; then
        add_result "10.1" "密码与传输安全" "TLS/加密协议版本" "fail" "检测到弱协议相关配置：$weak" "第10章" "禁用SSLv2/SSLv3/TLS1.0/1.1，仅启用TLS1.2及以上版本。"
    else
        add_result "10.1" "密码与传输安全" "TLS/加密协议版本" "manual" "OpenSSL版本：${opensslver:-未知}。未在sshd_config发现弱协议配置，具体业务系统（如Web服务器）TLS版本需人工核查证书与协议配置。" "第10章" "确保对外提供服务的TLS版本不低于1.2，使用符合国密/国际标准的强加密套件。"
    fi
}

# ============================================================
# 第5-9章 手动核查项（网络/物理/组织/制度/管理）
# ============================================================

add_manual_checks() {
    # ---- 第1章 补充 ----
    add_result "1.26_m" "系统安全" "日志留存与更新记录完整性" "manual" "系统更新日志、安全日志的完整留存情况需人工二次确认（自动化核查仅覆盖当前可读取部分）。" "第1章" "确保更新日志、安全事件日志按制度要求完整留存，不被提前清理。"

    # ---- 第2章 补充 ----
    add_result "2.3" "用户安全" "移动端应用锁屏与口令验证" "na" "本机为服务器/工作站操作系统，移动端应用锁屏及口令验证保护不适用于本检查对象。" "第2章" "如涉及移动端接入，应启用应用锁屏及口令验证保护。"
    add_result "2.11_m" "用户安全" "口令认证硬件Key（双因子）" "na" "是否采用硬件Key等双因子认证方式需结合具体业务系统评估，本次系统级核查不适用。" "第2章" "对高权限账户建议启用硬件Key等双因子认证方式。"

    # ---- 第3章 补充（数据安全）----
    add_result "3.2" "数据安全" "数据加密统一管控" "na" "用户/应用数据加密的统一管控措施属于业务系统层面，本次操作系统层核查不适用。" "第3章" "对敏感数据的获取、加密应实施统一管控。"
    add_result "3.3" "数据安全" "存储介质消磁处理" "na" "存储介质报废前消磁处理需结合资产管理流程人工核查，不适用于本次系统核查。" "第3章" "敏感存储介质报废前应采取消磁等不可恢复处理措施。"
    add_result "3.4" "数据安全" "重要数据全生命周期安全管理" "na" "重要数据传输、存储、销毁全过程安全管理需结合业务流程人工核查。" "第3章" "建立重要数据全生命周期（采集-传输-存储-销毁）安全管理机制。"
    add_result "3.6" "数据安全" "数据边界防泄漏(DLP)" "na" "数据边界防泄漏/防篡改措施属于业务/网络边界层面，本次系统核查不适用。" "第3章" "在数据边界部署防泄漏(DLP)与防篡改措施。"
    add_result "3.7" "数据安全" "数据存储访问日志与留存(≥180天)" "na" "数据存储系统的访问日志记录与180天留存需结合具体存储系统人工核查。" "第3章" "存储系统应记录访问日志并留存不少于180天。"
    add_result "3.9" "数据安全" "数据备份加密与隔离" "na" "数据备份是否采用加密、路径隔离、统一管控及操作日志记录需人工核查备份系统。" "第3章" "备份数据应加密存储，备份路径隔离并记录操作日志。"
    add_result "3.11" "数据安全" "数据访问权限最小化分级授权" "na" "数据访问权限的分级最小化授权需结合业务系统权限模型人工核查。" "第3章" "数据访问权限应按最小化原则分级授权。"
    add_result "3.12" "数据安全" "数据采集范围合规性" "na" "数据采集范围是否超出业务必要范围需结合业务需求人工核查。" "第3章" "数据采集范围应限定在业务必要范围内。"
    add_result "3.13" "数据安全" "备份介质加密" "na" "备份介质是否加密及必要性评估需人工核查备份管理制度。" "第3章" "视数据敏感级别对备份介质采取加密措施。"
    add_result "3.14" "数据安全" "敏感数据访问统一管控" "na" "敏感数据访问的统一管控与最小化授权需结合业务系统人工核查。" "第3章" "对敏感数据访问实施统一管控，权限最小化。"

    # ---- 第4章 补充（应用安全，业务层，非OS层）----
    add_result "4.1" "应用安全" "应用系统容灾措施" "na" "应用系统是否具备异地容灾/多活能力属于业务架构层面，本次系统核查不适用。" "第4章" "重要业务应用应具备容灾备份能力。"
    add_result "4.2" "应用安全" "应用系统补丁及时性" "na" "应用层（非OS层）补丁的及时安装情况需人工核查具体应用版本。" "第4章" "应用系统应及时安装厂商发布的安全补丁。"
    add_result "4.3" "应用安全" "开放端口应用身份认证与告警监测" "na" "对开放端口对应应用的身份认证与告警监测需结合业务系统人工核查。" "第4章" "对外开放端口的应用应具备身份认证及异常告警监测能力。"
    add_result "4.4" "应用安全" "敏感信息传输加密保护" "na" "业务层敏感信息获取过程的加密保护需人工核查具体应用实现。" "第4章" "业务应用获取敏感信息时应采取加密保护措施。"
    add_result "4.5" "应用安全" "互联网应用DDoS防御能力" "na" "对外提供服务的应用是否具备DDoS防御能力属于网络架构层面，不适用于本次系统核查。" "第4章" "面向互联网的应用应具备或依托上层设备的DDoS防御能力。"
    add_result "4.7" "应用安全" "网站静态化发布" "na" "网站应用是否采取静态页面发布方式需人工核查具体网站架构。" "第4章" "有条件的网站建议采取静态页面发布方式降低攻击面。"
    add_result "4.9" "应用安全" "Web应用防SQL注入/XSS攻击能力" "na" "Web应用自身的防注入/防XSS能力需人工核查代码实现或WAF防护效果。" "第4章" "Web应用应具备或依托WAF防SQL注入、XSS等攻击的能力。"
    add_result "4.10" "应用安全" "执行代码有效性校验" "na" "应用对上传/执行代码的有效性校验需人工核查具体代码实现。" "第4章" "应用应对可执行代码进行有效性与合法性校验。"
    add_result "4.11" "应用安全" "用户权限访问控制开发规范" "na" "基于用户权限的访问控制开发规范需人工核查应用设计文档。" "第4章" "应用开发应遵循最小权限访问控制规范。"
    add_result "4.13" "应用安全" "业务管理终端专人专用" "na" "业务管理终端是否专人专用需现场人工核查。" "第4章" "业务管理终端应指定专人专用并登记备案。"
    add_result "4.14" "应用安全" "基于用户角色的细粒度访问控制" "na" "基于用户角色授权的细粒度访问控制需人工核查应用RBAC实现。" "第4章" "应用应基于用户角色实现细粒度的访问控制。"
    add_result "4.19" "应用安全" "非通用协议接口通信" "na" "是否使用非通用协议/接口进行关键通信需人工核查具体业务实现。" "第4章" "关键业务通信建议避免使用完全通用、易被利用的协议接口。"
    add_result "4.20" "应用安全" "口令以外的认证方式" "na" "是否具备口令认证以外的认证方式（如证书、生物特征）需人工核查。" "第4章" "重要应用建议提供口令以外的认证方式。"
    add_result "4.25" "应用安全" "应用系统代码级安全漏洞（渗透测试）" "na" "应用系统是否存在代码级安全漏洞需通过渗透测试评估，不适用于本次系统核查。" "第4章" "定期开展渗透测试，修复代码级安全漏洞。"
    add_result "4.26" "应用安全" "账户异常登录/异常通信监测" "na" "账户异常登录、异常通信、异常文件上传等监测能力需人工核查具体应用/态势感知系统。" "第4章" "应用应具备账户异常行为监测能力，或依托态势感知平台。"
    add_result "4.27" "应用安全" "应用层攻击防护效果验证" "na" "抵御SQL注入/XSS/DoS等应用层攻击的实际效果需通过测试验证，不适用于本次系统核查。" "第4章" "定期验证应用层攻击防护措施的实际效果。"
    add_result "4.28" "应用安全" "统一身份认证(IAM)" "na" "是否接入统一身份认证平台需人工核查具体业务系统集成情况。" "第4章" "重要业务系统建议接入统一身份认证(IAM)平台。"
    add_result "4.30" "应用安全" "并发会话/连接数限制（应用层）" "na" "应用层对并发会话数、连接数、单用户连接数的限制需人工核查具体应用配置。" "第4章" "应用层应对并发会话及连接数进行合理限制。"

    # ---- 第5章 网络安全（网络设备/架构层，不适用于单机系统核查）----
    add_result "5.1" "网络安全" "网络拓扑图" "na" "网络拓扑图的绘制与维护需人工核查网络管理文档，不适用于单机系统核查。" "第5章" "绘制并维护最新网络拓扑图。"
    add_result "5.2" "网络安全" "专用网络设备部署" "na" "是否部署专用边界防护设备（防火墙/IDS/IPS等）需现场核查网络架构。" "第5章" "在网络边界部署专用安全防护设备。"
    add_result "5.3" "网络安全" "网络架构最小化原则" "na" "网络架构是否遵循最小化原则需人工核查网络设计文档。" "第5章" "网络架构设计应遵循最小化、按需开放原则。"
    add_result "5.4" "网络安全" "边界防护完整性" "na" "网络边界防护措施的完整性需现场核查所有出入口。" "第5章" "确保所有网络边界出入口均有防护措施覆盖。"
    add_result "5.5" "网络安全" "VLAN/安全域划分" "na" "VLAN划分与安全域隔离情况需人工登录网络设备核查。" "第5章" "按业务重要性划分VLAN及安全域，实现隔离。"
    add_result "5.6" "网络安全" "网络设备唯一标识" "na" "网络设备是否具备唯一标识管理需人工核查资产台账。" "第5章" "为网络设备分配唯一标识并纳入资产管理。"
    add_result "5.7" "网络安全" "远程管理加密" "na" "网络设备远程管理是否采用加密方式（SSH/HTTPS）需人工登录核查。" "第5章" "网络设备远程管理应使用加密协议，禁用Telnet/HTTP。"
    add_result "5.8" "网络安全" "网络设备厂商多样性" "na" "关键网络设备厂商是否存在单点依赖需人工核查采购记录。" "第5章" "关键网络设备建议避免厂商单一依赖。"
    add_result "5.9" "网络安全" "多播流量控制" "na" "多播流量的控制策略需人工登录网络设备核查。" "第5章" "对多播流量实施必要的控制策略。"
    add_result "5.10" "网络安全" "备份网络设备" "na" "关键网络设备是否具备热备/冷备需现场核查。" "第5章" "关键网络设备应配置备份设备，避免单点故障。"
    add_result "5.11" "网络安全" "安全运营中心(SOC)集成" "na" "网络设备日志是否已接入SOC/态势感知平台需人工核查集成情况。" "第5章" "网络设备安全事件应接入统一安全运营平台。"
    add_result "5.12" "网络安全" "单位间边界防护" "na" "不同单位/部门间网络边界的防护措施需人工核查网络架构。" "第5章" "在不同单位/部门网络边界部署防护措施。"
    add_result "5.13" "网络安全" "端口访问控制列表(ACL)" "na" "网络设备端口ACL配置需人工登录设备核查。" "第5章" "在网络设备端口配置访问控制列表限制非法访问。"
    add_result "5.14" "网络安全" "VPN接入安全" "na" "VPN远程接入的安全策略需人工核查VPN网关配置。" "第5章" "远程接入应通过VPN并采取强身份认证。"
    add_result "5.15" "网络安全" "802.1x接入认证" "na" "是否启用802.1x等接入认证机制需人工登录接入设备核查。" "第5章" "内网接入建议启用802.1x等认证机制。"
    add_result "5.16" "网络安全" "用户网络隔离" "na" "不同用户/部门间的网络隔离情况需人工核查VLAN/ACL配置。" "第5章" "不同用户群体间应实现必要的网络隔离。"
    add_result "5.17" "网络安全" "网络日志留存(≥180天)" "na" "网络设备日志留存时长需人工核查日志服务器配置，不适用于单机核查。" "第5章" "网络设备日志应留存不少于180天。"
    add_result "5.18" "网络安全" "IDS/IPS部署" "na" "入侵检测/防御系统的部署与规则更新需人工现场核查。" "第5章" "在网络关键节点部署并维护IDS/IPS。"
    add_result "5.19" "网络安全" "存储网络访问控制" "na" "存储网络(SAN/NAS)的访问控制策略需人工登录存储设备核查。" "第5章" "对存储网络实施严格的访问控制。"
    add_result "5.20" "网络安全" "网络设备冗余" "na" "核心网络设备的冗余部署情况需现场核查。" "第5章" "核心网络设备应采用冗余部署避免单点故障。"
    add_result "5.21" "网络安全" "网络设备口令策略" "na" "网络设备管理口令的复杂度与更新策略需人工登录设备核查。" "第5章" "网络设备管理口令应满足复杂度要求并定期更换。"
    add_result "5.22" "网络安全" "无线网络安全" "na" "无线接入点的加密方式与准入控制需现场核查。" "第5章" "无线网络应采用WPA2/WPA3加密并实施准入控制。"
    add_result "5.23" "网络安全" "网络安全域间访问审计" "na" "跨安全域访问的审计记录需人工核查网络审计系统。" "第5章" "对跨安全域的访问行为进行审计记录。"

    # ---- 第6章 物理安全（现场核查）----
    add_result "6.1" "物理安全" "机房门禁管理" "manual" "机房门禁系统的部署与出入记录管理需现场人工核查。" "第6章" "机房应部署门禁系统并留存出入记录。"
    add_result "6.2" "物理安全" "机房消防设施" "manual" "机房消防设施（气体灭火、烟感等）配置需现场人工核查。" "第6章" "机房应配置符合规范的消防设施。"
    add_result "6.3" "物理安全" "温湿度控制(GB2887-2011)" "manual" "机房温湿度是否符合GB2887-2011要求需现场人工核查监测记录。" "第6章" "机房温湿度控制应符合GB2887-2011标准要求。"
    add_result "6.4" "物理安全" "视频监控覆盖" "manual" "机房摄像录像覆盖范围及留存时长需现场人工核查。" "第6章" "机房重要区域应实现视频监控覆盖并留存记录。"
    add_result "6.5" "物理安全" "区域划分管理" "manual" "机房功能区域划分（办公区/机房区/存储区）需现场人工核查。" "第6章" "机房应实现功能区域划分并分级管控出入权限。"
    add_result "6.6" "物理安全" "存储介质管控" "manual" "存储介质（磁盘/磁带/U盘）的出入库管理需人工核查台账记录。" "第6章" "存储介质应建立出入库登记与管控台账。"
    add_result "6.7" "物理安全" "国产可控设备采购台账" "manual" "关键设备是否采用国产自主可控产品需人工核查采购台账。" "第6章" "关键设备应优先采购国产自主可控产品并建立台账。"
    add_result "6.8" "物理安全" "设备安全等级标签" "manual" "设备安全等级标签（色系标签）张贴情况需现场人工核查。" "第6章" "按设备安全等级张贴对应色系标签便于识别管理。"
    add_result "6.9" "物理安全" "设备防盗防破坏" "manual" "设备固定、防盗、防破坏措施需现场人工核查。" "第6章" "对关键设备采取防盗、防破坏固定措施。"

    # ---- 第7章 组织安全（文档/人员核查）----
    add_result "7.1" "组织安全" "安全组织架构" "manual" "信息安全组织架构是否健全需人工核查组织文件。" "第7章" "建立健全的信息安全组织架构。"
    add_result "7.2" "组织安全" "专职安全架构师配备" "manual" "是否配备专职安全架构师需人工核查人员编制记录。" "第7章" "应配备专职安全架构师负责整体安全规划。"
    add_result "7.3" "组织安全" "安全领导小组" "manual" "是否成立网络安全工作领导小组需人工核查文件与会议记录。" "第7章" "应成立由主要负责人牵头的网络安全工作领导小组。"
    add_result "7.4" "组织安全" "专职安全技术人员配备" "manual" "专职安全技术人员的配备情况需人工核查人员编制记录。" "第7章" "应配备与业务规模匹配的专职安全技术人员。"
    add_result "7.5" "组织安全" "应急响应体系" "manual" "网络安全应急响应组织体系是否建立需人工核查文档。" "第7章" "建立覆盖事前-事中-事后的应急响应组织体系。"

    # ---- 第8章 制度制定（文档核查）----
    add_result "8.1" "制度安全" "制度评审与发布机制" "manual" "安全管理制度的定期评审与发布机制需人工核查文档版本记录。" "第8章" "建立安全管理制度定期评审与正式发布机制。"
    add_result "8.2" "制度安全" "变更管理制度" "manual" "系统/配置变更管理制度的执行情况需人工核查变更记录。" "第8章" "建立并严格执行变更管理制度。"
    add_result "8.3" "制度安全" "备份策略制度" "manual" "数据备份策略（周期、留存、异地）是否制度化需人工核查文档。" "第8章" "制定并落实数据备份策略制度。"
    add_result "8.4" "制度安全" "制度审查机制" "manual" "安全管理制度的合规性审查机制需人工核查审查记录。" "第8章" "建立安全管理制度的定期合规性审查机制。"

    # ---- 第9章 管理实施（执行记录核查）----
    add_result "9.1" "管理安全" "应急演练" "manual" "网络安全应急预案的实际演练记录需人工核查演练台账。" "第9章" "定期组织网络安全应急演练并留存记录。"
    add_result "9.2" "管理安全" "统一管理平台" "manual" "设备/账户/日志是否接入统一管理平台需人工核查平台集成情况。" "第9章" "推动设备、账户、日志纳入统一管理平台管控。"
    add_result "9.3" "管理安全" "培训文档与记录" "manual" "安全意识与技能培训的文档及参与记录需人工核查。" "第9章" "定期开展安全培训并留存文档与参与记录。"
    add_result "9.4" "管理安全" "日常巡检制度" "manual" "系统日常安全巡检的执行与记录情况需人工核查巡检台账。" "第9章" "建立并落实系统日常安全巡检制度。"

    # ---- 第10章 补充 ----
    add_result "10.1_m" "密码与传输安全" "密钥签发与证书管控" "na" "密钥协商认证、证书签发与管控体系需结合PKI/密钥管理系统人工核查，不适用于单机系统核查。" "第10章" "建立密钥/证书全生命周期管控体系（签发、更新、吊销）。"
}

# ============================================================
# 汇总与报告生成
# ============================================================

print_summary() {
    local pass=0 fail=0 manual=0 na=0
    local i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1)) ;;
            fail) fail=$((fail+1)) ;;
            manual) manual=$((manual+1)) ;;
            na) na=$((na+1)) ;;
        esac
    done
    echo "核查完成，共 $R_COUNT 项："
    echo "  合规(pass)：$pass    不合规(fail)：$fail    需人工核查(manual)：$manual    不适用(na)：$na"
}

html_esc() {
    local s="$1"
    s="${s//&/&amp;}"
    s="${s//</&lt;}"
    s="${s//>/&gt;}"
    printf '%s' "$s"
}

status_cn() {
    case "$1" in
        pass) echo "合规" ;;
        fail) echo "不合规" ;;
        manual) echo "需人工核查" ;;
        na) echo "不适用" ;;
        *) echo "$1" ;;
    esac
}

status_color() {
    case "$1" in
        pass) echo "#2e7d32" ;;
        fail) echo "#c62828" ;;
        manual) echo "#ef6c00" ;;
        na) echo "#757575" ;;
        *) echo "#000000" ;;
    esac
}

build_table_rows() {
    local i
    for ((i=1; i<=R_COUNT; i++)); do
        local color; color="$(status_color "${R_STATUS[$i]}")"
        local scn; scn="$(status_cn "${R_STATUS[$i]}")"
        cat <<ROW
<tr>
<td>$(html_esc "${R_CHAPTER[$i]}")</td>
<td>$(html_esc "${R_ID[$i]}")</td>
<td>$(html_esc "${R_CAT[$i]}")</td>
<td>$(html_esc "${R_TITLE[$i]}")</td>
<td style="color:$color;font-weight:bold;">$scn</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "${R_REC[$i]}")</td>
<td>$(html_esc "${R_GUIDE[$i]}")</td>
</tr>
ROW
    done
}

# 按状态筛选生成表格行，与Windows版GenerateHTML的分区表结构一致（未通过/需人工核查/不适用/通过）
build_table_rows_by_status() {
    local want="$1" i
    for ((i=1; i<=R_COUNT; i++)); do
        [ "${R_STATUS[$i]}" != "$want" ] && continue
        local color; color="$(status_color "${R_STATUS[$i]}")"
        local scn; scn="$(status_cn "${R_STATUS[$i]}")"
        cat <<ROW
<tr>
<td>$(html_esc "${R_CHAPTER[$i]}")</td>
<td>$(html_esc "${R_ID[$i]}")</td>
<td>$(html_esc "${R_CAT[$i]}")</td>
<td>$(html_esc "${R_TITLE[$i]}")</td>
<td style="color:$color;font-weight:bold;">$scn</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "${R_REC[$i]}")</td>
<td>$(html_esc "${R_GUIDE[$i]}")</td>
</tr>
ROW
    done
}

generate_html() {
    local outfile="$OUT_DIR/配置核查报告_${STAMP}.html"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1)) ;;
            fail) fail=$((fail+1)) ;;
            manual) manual=$((manual+1)) ;;
            na) na=$((na+1)) ;;
        esac
    done
    {
        cat <<HTMLHEAD
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>配置核查报告</title>
<style>
body{font-family:"Microsoft YaHei",Arial,sans-serif;margin:20px;color:#222;}
h1{color:#1a3c6e;}
.meta{background:#f0f4f8;padding:10px 15px;border-radius:6px;margin-bottom:15px;}
.summary{display:flex;gap:15px;margin-bottom:20px;}
.card{flex:1;padding:12px;border-radius:6px;text-align:center;color:#fff;}
.card.pass{background:#2e7d32;} .card.fail{background:#c62828;}
.card.manual{background:#ef6c00;} .card.na{background:#757575;}
table{border-collapse:collapse;width:100%;font-size:13px;}
th,td{border:1px solid #ccc;padding:6px 8px;vertical-align:top;}
th{background:#1a3c6e;color:#fff;position:sticky;top:0;}
tr:nth-child(even){background:#f7f9fb;}
</style>
</head>
<body>
<h1>配置核查报告 - 中标麒麟/银河麒麟版</h1>
<div class="meta">
<div>系统：$(html_esc "$OS_NAME $OS_VER")　（$(html_esc "$KYLIN_TYPE")）</div>
<div>主机名：$(html_esc "$HOSTNAME_STR")　内核：$(html_esc "$KERNEL_VER")</div>
<div>核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>
<div>参考标准：配置核查作业指导书正式版2026_4_1</div>
</div>
<div class="summary">
<div class="card pass">合规<br>$pass</div>
<div class="card fail">不合规<br>$fail</div>
<div class="card manual">需人工核查<br>$manual</div>
<div class="card na">不适用<br>$na</div>
</div>
HTMLHEAD
        if [ "$fail" -gt 0 ]; then
            echo "<h2>一、未通过（$fail 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status fail
            echo "</table>"
        fi
        if [ "$manual" -gt 0 ]; then
            echo "<h2>二、需人工核查（$manual 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status manual
            echo "</table>"
        fi
        if [ "$na" -gt 0 ]; then
            echo "<h2>三、不适用（$na 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status na
            echo "</table>"
        fi
        if [ "$pass" -gt 0 ]; then
            echo "<h2>四、通过（$pass 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status pass
            echo "</table>"
        fi
        cat <<HTMLFOOT
</body>
</html>
HTMLFOOT
    } > "$outfile"
    echo "HTML报告已生成：$outfile"
}

generate_xls() {
    local outfile="$OUT_DIR/配置核查报告_${STAMP}.xls"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1)) ;;
            fail) fail=$((fail+1)) ;;
            manual) manual=$((manual+1)) ;;
            na) na=$((na+1)) ;;
        esac
    done
    {
        cat <<XLSHEAD
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]>
<xml>
<x:ExcelWorkbook>
<x:ExcelWorksheets>
<x:ExcelWorksheet>
<x:Name>配置核查报告</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet>
</x:ExcelWorksheets>
</x:ExcelWorkbook>
</xml>
<![endif]-->
<style>
table{border-collapse:collapse;}
th,td{border:1px solid #999;padding:4px 6px;font-family:"Microsoft YaHei",Arial;font-size:12px;mso-number-format:"\@";}
th{background:#1a3c6e;color:#fff;font-weight:bold;}
</style>
</head>
<body>
<p>配置核查报告 - 中标麒麟/银河麒麟版　系统：$(html_esc "$OS_NAME $OS_VER")（$(html_esc "$KYLIN_TYPE")）　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>
XLSHEAD
        if [ "$fail" -gt 0 ]; then
            echo "<p><b>一、未通过（$fail 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status fail
            echo "</table>"
        fi
        if [ "$manual" -gt 0 ]; then
            echo "<p><b>二、需人工核查（$manual 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status manual
            echo "</table>"
        fi
        if [ "$na" -gt 0 ]; then
            echo "<p><b>三、不适用（$na 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status na
            echo "</table>"
        fi
        if [ "$pass" -gt 0 ]; then
            echo "<p><b>四、通过（$pass 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"
            build_table_rows_by_status pass
            echo "</table>"
        fi
        cat <<XLSFOOT
</body>
</html>
XLSFOOT
    } > "$outfile"
    echo "Excel(.xls)报告已生成：$outfile"
}

# ============================================================
# 主流程
# ============================================================

# ---------- 第1章 系统安全 ----------
check_1_1_patch
check_1_2_antivirus
check_1_3_services
check_1_4_firewall
check_1_5_selinux
check_1_6_ssh_crypto
check_1_17_redis
check_1_24_auditpolicy
check_1_25_defender
check_1_18_resourcelimits
check_1_26_securitylogs
check_1_7_dbaccounts
check_1_8_dbprocedures
check_1_9_dbgrants
check_1_10_dbaccess
check_1_11_dbbackup
check_1_12_dbauditplugin
check_1_13_dbisolation
check_1_14_middleware
check_1_15_dbremote
check_1_16_17_inputcheck
check_1_19_dbdefaults
check_1_20_dbpolicy
check_1_21_dbaudit
check_1_22_dbauditretention
check_1_23_dbappsep

# ---------- 第2章 用户安全 ----------
check_2_1_passwordexpiry
check_2_4_uniqueaccounts
check_2_5_redundantservices
check_2_2_userpatch
check_2_6_illegalinternet
check_2_8_endpointcontrol
check_2_15_securitypolicy
check_2_16_patchlatest
check_2_7_fstype
check_2_9_riskysoftware
check_2_10_usb
check_2_11_passwordpolicy
check_2_12_wireless
check_2_13_outboundcontrol
check_2_14_useraudit

# ---------- 第3章 数据安全 ----------
check_3_1_encryption
check_3_5_logprotection
check_3_10_shares
check_3_8_storagepartition

# ---------- 第4章 应用安全 ----------
check_4_15_nla
check_4_16_roleseparation
check_4_17_iprestrict
check_4_18_lockout
check_4_22_portseparation
check_4_23_appdbsep
check_4_24_appaudit
check_4_6_waf
check_4_8_antitamper
check_4_12_sessionlimit
check_4_21_defaultport
check_4_29_remotemgmt
check_4_31_appbackup
check_4_32_domesticsoftware

# ---------- 第10章 密码与传输安全 ----------
check_10_1_tls

# ---------- 第5-9章 手动核查项 ----------
add_manual_checks

# ---------- 生成报告 ----------
echo ""
echo "================================================================"
print_summary
generate_html
generate_xls

echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS 报告。"
