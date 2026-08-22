#!/bin/bash
# ============================================================
# 配置核查工具 - 中标麒麟/银河麒麟版（无需Python，纯Bash实现）
# 适用系统：中标麒麟 NeoKylin（yum/rpm，多基于CentOS/RHEL）
#           银河麒麟 Kylin OS（apt/dpkg，多基于Ubuntu/Debian）
# 参考标准：配置核查作业指导书v2.2
# 运行方式：sudo bash run.sh 或 sudo bash check_kylin.sh
# 输出文件：output/配置核查报告_日期时间.html /.xls
# ============================================================
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT_DIR="$SCRIPT_DIR/output"
mkdir -p "$OUT_DIR"
STAMP="$(date +%Y%m%d_%H%M%S)"

# 零依赖 .xlsx 生成器（可选，需要 zip 命令；缺失时自动降级为 .xls）
[ -f "$SCRIPT_DIR/lib_xlsx.sh" ] && source "$SCRIPT_DIR/lib_xlsx.sh"

# ---------- 结果存储（并行数组）----------
R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=()
R_COUNT=0

# 《配置核查作业指导书v2.2》目录页码与条款原文对照表（按核查编号查询）
# 用于在报告中提示测试人员应翻阅指导书哪一页、对照哪一条原文表述
guide_ref() {
    local chap="${1%%.*}" cname="" clause="" title=""
    case "$chap" in
        1) cname="第1章 系统安全" ;;
        2) cname="第2章 用户安全" ;;
        3) cname="第3章 数据安全" ;;
        4) cname="第4章 应用安全" ;;
        5) cname="第5章 网络安全" ;;
        6) cname="第6章 物理安全" ;;
        7) cname="第7章 组织机构" ;;
        8) cname="第8章 规章制度" ;;
        9) cname="第9章 管理实施" ;;
        10) cname="第10章 协议安全审计" ;;
        *) cname="" ;;
    esac
    [ -z "$cname" ] && { echo ""; return; }
    case "$1" in
        1.1) clause=1.1; title="操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.2) clause=1.2; title="操作系统应安装防病毒软件并及时升级" ;;
        1.3) clause=1.3; title="操作系统应按需求裁剪服务和端口" ;;
        1.4) clause=1.4; title="操作系统应具备防火墙功能" ;;
        1.5) clause=1.5; title="操作系统应停用冗余网络设置" ;;
        1.6) clause=1.6; title="操作系统远程管理应开放唯一管理服务，指定管理终端并采取传输加密保护措施" ;;
        1.7) clause=1.7; title="数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" ;;
        1.8) clause=1.8; title="数据库管理系统应删除冗余存储过程" ;;
        1.9) clause=1.9; title="数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" ;;
        1.10) clause=1.10; title="数据库管理系统应具有自主访问控制功能" ;;
        1.11) clause=1.11; title="数据库管理系统应具有备份和恢复功能" ;;
        1.12) clause=1.12; title="数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.13) clause=1.13; title="数据库管理系统的数据应和其它应用的数据分类独立存储" ;;
        1.14) clause=1.14; title="中间件应采取限制运行权限和使用安全管理通道等安全加固措施" ;;
        1.15) clause=1.15; title="应具备数据库管理系统超级管理员远程登录限制能力" ;;
        1.16) clause=1.16; title="应具备数据库管理系统输入（参数）检查能力" ;;
        1.16-17) clause="1.16、1.17"; title="应具备数据库管理系统输入（参数）检查能力 / 应限制操作系统开放的远程管理服务或端口" ;;
        1.17) clause=1.17; title="应限制操作系统开放的远程管理服务或端口" ;;
        1.18) clause=1.18; title="应限制用户对服务器资源的最大或最小使用限度" ;;
        1.19) clause=1.19; title="应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.20) clause=1.20; title="数据库管理系统应配置安全策略" ;;
        1.21) clause=1.21; title="数据库管理系统应具有行级或列级审计功能" ;;
        1.22) clause=1.22; title="数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) clause=1.23; title="数据库管理系统仅为应用服务器提供访问服务" ;;
        1.24) clause=1.24; title="应具备日志审计能力，审计日志至少保留180天" ;;
        1.25) clause=1.25; title="检查是否具备边界保护能力，是否可以抗攻击、防篡改" ;;
        1.26|1.26_m) clause=1.26; title="检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" ;;
        1.27) clause=1.27; title="操作系统、数据库管理系统、中间件、办公软件等基础软件应采用具有完备的售后技术支持与服务的正版或定制软件" ;;
        2.1) clause=2.1; title="服务器和用户计算机应设置登录口令" ;;
        2.2) clause=2.2; title="用户计算机应根据需要安装补丁程序" ;;
        2.3) clause=2.3; title="用户应设置用户应用口令，通过认证后使用信息服务" ;;
        2.4) clause=2.4; title="用户计算机应具有互不相同的用户名和口令" ;;
        2.5) clause=2.5; title="用户计算机应关闭冗余系统服务和端口" ;;
        2.6) clause=2.6; title="用户计算机应具备阻断和告警非法连接互联网的能力" ;;
        2.7) clause=2.7; title="用户计算机应具有文件保护功能" ;;
        2.8) clause=2.8; title="应采取终端管控措施（统一配置、安全加固、网络访问控制、外设接口管控、软件进程管控、无线模块禁用等）" ;;
        2.9) clause=2.9; title="用户计算机应禁止安装与工作无关的软件" ;;
        2.10) clause=2.10; title="用户计算机USB接口应禁止私自连接对拷线和手机、媒体播放设备等个人移动电子设备" ;;
        2.11|2.11_m) clause=2.11; title="用户计算机登录应使用基于专用物理部件或生物特征的多因素身份认证方式，应设置超时锁屏（不超过5min），口令不少于10个字符，更换周期不超过30d" ;;
        2.12) clause=2.12; title="用户计算机应物理拆除Wi-Fi、红外、蓝牙等无线模块" ;;
        2.13) clause=2.13; title="用户计算机应采取非法外联阻断、文件输出管控等措施" ;;
        2.14) clause=2.14; title="应具备用户行为审计能力；审计日志应至少保留xx天" ;;
        2.15) clause=2.15; title="检查安全策略配置情况和设置功能" ;;
        2.16) clause=2.16; title="检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" ;;
        3.1) clause=3.1; title="集中存储的涉密数据应采取加密保护措施" ;;
        3.2) clause=3.2; title="用户计算机之间的数据交互、文件传输等应统一管控和审计" ;;
        3.3) clause=3.3; title="涉密存储载体在降密级使用前或重大演训活动结束后，应采取数据写覆盖方法及时清除数据" ;;
        3.4) clause=3.4; title="对确定销毁的涉密载体，应采取消磁、粉碎、溶解、化浆和熔化等方法进行销毁" ;;
        3.5) clause=3.5; title="网络、系统、应用和用户行为等日志应采取读写控制、加密、变换、完整性校验等保护措施" ;;
        3.6) clause=3.6; title="网络边界应具备信息过滤、敏感内容识别等数据防泄漏能力" ;;
        3.7) clause=3.7; title="数据存储系统管理登录至少采取验证码等增强措施，修改默认设置，审计日志至少保留180天" ;;
        3.8) clause=3.8; title="数据存储系统应根据重要程度划分不同存储区块，并设置用户访问权限" ;;
        3.9) clause=3.9; title="检查数据传输过程是否按要求加密，传输路径是否合理，是否统一管控、留有日志记录、具有防泄漏措施" ;;
        3.10) clause=3.10; title="检查数据共享是否合理，是否存在安全隐患" ;;
        3.11) clause=3.11; title="检查对数据的访问是否按照权限分级访问，是否对访问行为进行审计" ;;
        3.12) clause=3.12; title="检查数据采集是否超出业务需求范围" ;;
        3.13) clause=3.13; title="检查数据各环节处理是否满足密级相应的保密要求" ;;
        3.14) clause=3.14; title="检查对数据的访问是否按照权限分级访问，对大数据的访问是否提供统一管控和访问控制" ;;
        4.1) clause=4.1; title="应用系统软件应采取备份措施" ;;
        4.2) clause=4.2; title="应用系统软件应及时安装补丁程序，且更新所用升级包应经过安全性测试" ;;
        4.3) clause=4.3; title="基于可信根对应用系统软件进行可信验证，可信性受到破坏后报警" ;;
        4.4) clause=4.4; title="提供公共信息服务的服务器应与涉密信息服务器分设，专用服务器应只提供专用服务" ;;
        4.5) clause=4.5; title="提供公共信息服务的服务器应具备防DDoS攻击能力" ;;
        4.6) clause=4.6; title="Web应用系统服务器应采取Web防护措施" ;;
        4.7) clause=4.7; title="网站服务宜以静态页面形式发布" ;;
        4.8) clause=4.8; title="网站应采取网页防篡改措施，防止对信息内容的非法修改" ;;
        4.9) clause=4.9; title="Web应用系统应具备防范SQL注入、跨站脚本等攻击能力" ;;
        4.10) clause=4.10; title="Web应用系统应具备执行代码有效验证能力" ;;
        4.11) clause=4.11; title="应具备用户授权访问控制能力" ;;
        4.12) clause=4.12; title="应具备访问应用最大并发会话连接数限制能力" ;;
        4.13) clause=4.13; title="业务管理终端专设专用" ;;
        4.14) clause=4.14; title="应具有基于用户角色的授权访问控制能力，主体细粒度达到用户级或进程级，客体细粒度达到文件级、表和记录级、字段级" ;;
        4.15) clause=4.15; title="文电等文档专用业务处理系统应具有签名验证、密级标识等功能" ;;
        4.16) clause=4.16; title="远程管理应采取加密保护措施" ;;
        4.17) clause=4.17; title="应支持管理员、安全员、审计员三权分立的职责划分，禁止设立超级管理员" ;;
        4.18) clause=4.18; title="业务管理终端登录应采取网络地址限制措施" ;;
        4.19) clause=4.19; title="应具有结束会话、限定登录错误次数和自动退出等登录失败处理功能" ;;
        4.20) clause=4.20; title="宜使用自主设计开发的网络服务、协议、接口等，增强应用安全" ;;
        4.21) clause=4.21; title="应具有基于专用物理部件或生物特征多因素的、与JD密码算法相结合的数字证书用户身份认证功能" ;;
        4.22) clause=4.22; title="应更改Web应用系统默认服务发布端口" ;;
        4.23) clause=4.23; title="应分开设置管理端口与应用端口" ;;
        4.24) clause=4.24; title="应用服务和数据存储应部署在不同的服务器上" ;;
        4.25) clause=4.25; title="应具有对所有访问行为和管理行为的日志审计功能，审计日志应至少保留180d" ;;
        4.26) clause=4.26; title="应经过代码级安全漏洞挖掘" ;;
        4.27) clause=4.27; title="检查是否具备对人机接口输入、网络通信输入、文件输入的数据进行格式和长度检查的功能" ;;
        4.28) clause=4.28; title="检查是否能够有效检测并防御SQL注入、网页篡改、跨站脚本、拒绝服务等应用层攻击" ;;
        4.29) clause=4.29; title="检查是否具有用户访问权限统一管理功能" ;;
        4.30) clause=4.30; title="检查是否能够设置统一管理措施，是否对远程管理进行限制" ;;
        4.31) clause=4.31; title="检查是否能够设置最大并发会话连接数、会话建立速率、单用户并发会话数" ;;
        4.32) clause=4.32; title="检查所有应用是否具备备份与恢复功能（指导书原文4.33为国产化适配条款，与本项编号存在1位偏差，请人工核对）" ;;
        5.*) clause=""; title="指导书本章原文未细分小节编号，以下网络设备类核查项供参考，均标记不适用" ;;
        6.1) clause=6.1; title="使用的网络设备、服务器、终端等，应选用进入《全军计算机及网络设备集中采购目录》的产品" ;;
        6.2) clause=6.2; title="使用的安全网关、防火墙等信息安全产品，应通过信息安全测评认证机构的认证" ;;
        6.3) clause=6.3; title="机房应有序、规范、合理走线布线，明确互联网区域和内部区域，粘贴标识，区分不同线路" ;;
        6.4) clause=6.4; title="机房应符合GB 2887-2011中4.6.1所规定的温湿度等要求" ;;
        6.5) clause=6.5; title="机房应安装安防监控设备，对机房的人员进出、设备操作等进行监管" ;;
        6.6) clause=6.6; title="机房应将涉密区域和互联网区域设置于不同场所" ;;
        6.7) clause=6.7; title="在涉密网络中使用过的打印机、复印机、刻录机、扫描仪、存储载体等设备严禁在互联网中使用" ;;
        6.8) clause=6.8; title="应选用《关键软硬件自主可控产品名录》中的芯片类、计算机及外设类、网络设备类、安全防护设备类、存储设备类等产品" ;;
        6.9) clause=6.9; title="与安全防护等级低的网络使用不同色系线缆进行严格隔离区分" ;;
        7.1) clause=7.1; title="应设置安全管理机构，保证系统安全措施的落实" ;;
        7.2) clause=7.2; title="应配备专职系统安全保密管理人员，负责系统安全措施的落实" ;;
        7.3) clause=7.3; title="应具有安全管理领导机构，督导系统安全措施的落实" ;;
        7.4) clause=7.4; title="应具有系统安全保密技术管理人员，具体负责安全保密技术措施的落实" ;;
        7.5) clause=7.5; title="应具有完善的应急响应体系，应对突发事件" ;;
        8.1) clause=8.1; title="应具有日常安全管理制度、入网审批制度和系统安全保密检查制度等" ;;
        8.2) clause=8.2; title="应具有安全管理操作规程、安全监控操作规程、安全审计操作规程、应急响应操作规程等" ;;
        8.3) clause=8.3; title="应具有日常系统备份制度和存储载体使用管理制度" ;;
        8.4) clause=8.4; title="应具有脆弱性分析等规程" ;;
        9.1) clause=9.1; title="应具有应急响应预案，发生危及系统安全的事件时应根据操作规程及时采取措施" ;;
        9.2) clause=9.2; title="安全保密管理人员应能对网络安全防护设备进行统一配置、管理" ;;
        9.3) clause=9.3; title="应具有与实际情况相符且完整的安全保密策略文档和安全保密技术及产品配置的详细记录" ;;
        9.4) clause=9.4; title="安全保密技术管理人员应对每日网络和系统运行情况实施安全审计，每月组织安全性检测" ;;
        10.1|10.1_m) clause=""; title="指导书本章原文未细分小节编号，审计协议的机密性、完整性、认可性、不可否认性" ;;
        *) clause=""; title="" ;;
    esac
    if [ -z "$title" ]; then
        echo "《配置核查作业指导书》${cname}"
    elif [ -n "$clause" ]; then
        echo "《配置核查作业指导书》${cname} ${clause}：${title}"
    else
        echo "《配置核查作业指导书》${cname}（${title}）"
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
echo "  参考标准：配置核查作业指导书v2.2"
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
    local mtime=""
    if [ "$PKG_MGR" = "apt" ]; then
        mtime="$(stat -c %Y /var/lib/apt/periodic/update-success-stamp 2>/dev/null)"
        [ -z "$mtime" ] && mtime="$(stat -c %Y /var/log/apt/history.log 2>/dev/null)"
    else
        mtime="$(stat -c %Y /var/lib/rpm/Packages 2>/dev/null)"
        [ -z "$mtime" ] && mtime="$(stat -c %Y /var/log/yum.log 2>/dev/null)"
    fi
    local days=""
    if [ -n "$mtime" ]; then
        days=$(( ($(date +%s) - mtime) / 86400 ))
    fi
    if [ "$PKG_MGR" = "apt" ]; then
        local cnt
        cnt="$(run_cmd "apt list --upgradable 2>/dev/null | grep -v '^Listing' | wc -l")"
        cnt="${cnt:-0}"
        if [ "$cnt" -gt 0 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "fail" "检测到 $cnt 个可升级的软件包（apt list --upgradable），系统补丁未安装到最新。" "第1章" "使用 apt update && apt upgrade 及时安装安全补丁。"
        elif [ -n "$days" ] && [ "$days" -le 90 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "pass" "apt 未发现可升级软件包，且距上次系统更新 $days 天（90天内），补丁状态较新。" "第1章" "定期检查并安装系统安全补丁。"
        else
            add_result "1.1" "系统安全" "补丁安装情况" "manual" "apt 未发现可升级软件包，但无法确认系统近期是否执行过更新（距上次更新${days:-未知}天或无法读取更新记录/离线），需人工确认补丁现状。" "第1章" "确认系统更新源可达性并核实补丁是否安装到最新。"
        fi
    else
        local out
        out="$(run_cmd "yum check-update 2>/dev/null | grep -Ev '^$|^Loaded|^Last metadata|^Obsoleting|^Security:' | wc -l")"
        out="${out:-0}"
        if [ "$out" -gt 0 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "fail" "yum check-update 检测到约 $out 个可更新软件包，系统补丁未安装到最新。" "第1章" "使用 yum update 及时安装安全补丁。"
        elif [ -n "$days" ] && [ "$days" -le 90 ] 2>/dev/null; then
            add_result "1.1" "系统安全" "补丁安装情况" "pass" "yum 未发现可更新软件包，且距上次系统更新 $days 天（90天内），补丁状态较新。" "第1章" "定期检查并安装系统安全补丁。"
        else
            add_result "1.1" "系统安全" "补丁安装情况" "manual" "yum 未发现可更新软件包，但无法确认系统近期是否执行过更新（距上次更新${days:-未知}天或无法读取更新记录/离线），需人工确认补丁现状。" "第1章" "确认系统更新源可达性并核实补丁是否安装到最新。"
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
        add_result "1.3" "系统安全" "高危服务/端口" "manual" "未检测到 telnet/rsh/rlogin/tftp/rexec 等已知高危服务，但系统运行的服务与开放端口是否均为业务必需，需人工逐一确认。" "第1章" "核查运行服务列表与监听端口，停用并禁用非业务必需的服务。"
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
        local weak strong
        weak="$(grep -Ei '^\s*Ciphers' "$cfg" | grep -Ei '3des|arcfour|rc4|blowfish|cbc')"
        strong="$(grep -Ei '^\s*Ciphers' "$cfg" | grep -Eiv '3des|arcfour|rc4|blowfish|cbc')"
        if [ -n "$weak" ]; then
            add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "fail" "sshd_config 中配置了弱加密算法：$weak" "第1章" "禁用 3DES/RC4/CBC 等弱加密套件，使用 AES-GCM/ChaCha20 等强算法。"
        elif [ -n "$strong" ]; then
            add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "pass" "sshd_config 显式配置了加密算法且未含已知弱算法：$strong" "第1章" "定期核查 SSH 加密套件配置，避免使用弱算法。"
        else
            add_result "1.6" "系统安全" "远程管理加密强度（SSH）" "manual" "sshd_config 未显式配置 Ciphers，实际加密强度取决于系统默认套件，需人工确认默认套件不含弱算法。" "第1章" "建议显式配置 Ciphers 为 AES-GCM/ChaCha20 等强算法，避免依赖默认值。"
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

check_1_27_softwarelicensing() {
    local osinfo licinfo findings=""
    osinfo="$(run_cmd "cat /etc/os-release 2>/dev/null | grep -E '^(NAME|VERSION)=' | tr '\n' ' '")"
    [ -n "$osinfo" ] && findings="${findings}  · 系统：${osinfo}"
    licinfo="$(run_cmd "cat /etc/.kyinfo 2>/dev/null | head -5 | tr '\n' ' '")"
    if [ -n "$licinfo" ]; then
        findings="${findings}  · 授权信息：${licinfo}"
    else
        findings="${findings}  · 未找到 /etc/.kyinfo 授权文件"
    fi
    local repo
    repo="$(run_cmd "cat /etc/yum.repos.d/NeoKylin.repo 2>/dev/null | grep -E '^baseurl|^mirrorlist' | head -3 | tr '\n' ' '")"
    [ -n "$repo" ] && findings="${findings}  · YUM 源：${repo}"
    add_result "1.27" "系统安全" "基础软件正版/授权及售后技术支持" "manual" "操作系统版本与授权信息已自动核查，但正版/定制软件及售后技术支持需人工核对采购与授权凭证。${findings}" "第1章" "核对操作系统、数据库、中间件、办公软件的采购合同、正版授权证书及售后技术支持合同。"
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
    elif [ -n "$listen" ]; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "数据库监听的地址非公网暴露（0.0.0.0）：$listen。绑定配置：${bind:-默认}" "第1章" "持续保持数据库仅对可信来源开放访问。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "manual" "检测到数据库服务($DB_TYPE)，但未在标准端口(3306/5236/54321)检测到监听，可能使用自定义端口，需人工确认监听地址与访问控制。" "第1章" "确认数据库实际监听端口与地址，限制其仅对可信来源开放。"
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

check_2_3_apppass() {
    # 指导书 2.3：用户应用口令（OS 层信号：空口令账户 / 图形界面自动登录）
    if [ ! -r /etc/shadow ]; then
        add_result "2.3" "用户安全" "应用口令与认证登录" "manual" "无法读取 /etc/shadow，请以 root/sudo 运行后核查空口令账户与自动登录配置。" "第2章" "确保所有账户设置口令并通过认证后使用信息服务。"
        return
    fi
    local empty_users autologin
    empty_users="$(awk -F: '($2 == "") {print $1}' /etc/shadow 2>/dev/null | head -5 | tr '\n' ' ')"
    autologin=""
    if ls /etc/lightdm/*.conf >/dev/null 2>&1; then
        autologin="$(grep -hEi '^[[:space:]]*autologin-(user|session)[[:space:]]*=[[:space:]]*[^[:space:]#]+' /etc/lightdm/*.conf 2>/dev/null | head -2)"
    fi
    [ -z "$autologin" ] && [ -f /etc/sddm.conf ] && autologin="$(grep -Ei '^[[:space:]]*autologin' /etc/sddm.conf 2>/dev/null | grep -vi '^[[:space:]]*#' | head -2)"
    [ -z "$autologin" ] && ls /etc/gdm*/custom.conf >/dev/null 2>&1 && autologin="$(grep -hEi '^[[:space:]]*AutomaticLoginEnable(d)?[[:space:]]*=[[:space:]]*true' /etc/gdm*/custom.conf 2>/dev/null | head -2)"
    if [ -n "$empty_users" ]; then
        add_result "2.3" "用户安全" "应用口令与认证登录" "fail" "/etc/shadow 发现空口令账户：${empty_users}，可无口令登录。" "第2章" "为空口令账户设置口令：passwd 用户名。"
    elif [ -n "$autologin" ]; then
        add_result "2.3" "用户安全" "应用口令与认证登录" "fail" "图形界面启用了自动登录（${autologin}），登录无需口令认证。" "第2章" "关闭 lightdm/sddm/gdm 的 autologin/AutomaticLogin 配置。"
    else
        add_result "2.3" "用户安全" "应用口令与认证登录" "pass" "未发现空口令账户与自动登录配置，账户均需口令认证。" "第2章" "应用层接入同样启用口令认证，杜绝绕过认证使用信息服务。"
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
    case "$fstype" in
        ext3|ext4|xfs|btrfs)
            add_result "2.7" "用户安全" "文件系统安全" "pass" "根分区文件系统为 ${fstype}，原生支持权限位与 ACL，具备文件保护能力。" "第2章" "持续通过权限位/ACL 设置关键目录访问控制。"
            ;;
        *)
            add_result "2.7" "用户安全" "文件系统安全" "fail" "根分区文件系统为 ${fstype:-未知}，不支持或未确认支持权限位/ACL。" "第2章" "关键分区应采用 ext4/xfs 等支持权限位/ACL 的文件系统。"
            ;;
    esac
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
        add_result "2.9" "用户安全" "高风险软件" "manual" "未检测到已知高风险远程控制软件（teamviewer/anydesk/todesk/sunlogin），但可能存在其他名称的同类软件，需人工核查已安装软件清单。" "第2章" "核查已安装软件清单，卸载未经授权的远程控制等高危软件。"
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
            local minlen_num
            minlen_num="$(echo "$minlen" | grep -oE '[0-9]+' | head -1)"
            if [ -n "$minlen_num" ] && [ "$minlen_num" -ge 8 ] 2>/dev/null; then
                add_result "2.11" "用户安全" "口令复杂度策略" "pass" "$pamfile 中配置了密码复杂度策略：$minlen（最小长度 $minlen_num 位，≥8）。" "第2章" "确保密码包含大小写字母、数字、特殊字符中至少3类。"
            elif [ -n "$minlen_num" ]; then
                add_result "2.11" "用户安全" "口令复杂度策略" "fail" "$pamfile 中 minlen=$minlen_num，密码最小长度不足8位，强度不够。" "第2章" "将 pam_pwquality 的 minlen 调整为不少于8。"
            else
                add_result "2.11" "用户安全" "口令复杂度策略" "manual" "$pamfile 中配置了 $minlen 但无法解析长度数值，需人工确认复杂度要求是否满足。" "第2章" "配置PAM密码复杂度策略（长度≥8、字符种类、历史）。"
            fi
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
    local out rules drop
    out="$(run_cmd "iptables -S OUTPUT 2>/dev/null")"
    rules="$(echo "$out" | grep -c .)"
    drop="$(echo "$out" | grep -Ei 'DROP|REJECT' | wc -l)"
    if [ "${rules:-0}" -gt 1 ] 2>/dev/null && [ "${drop:-0}" -gt 0 ] 2>/dev/null; then
        add_result "2.13" "用户安全" "外联控制" "pass" "iptables OUTPUT 链已配置 $rules 条规则，其中含 $drop 条拒绝(DROP/REJECT)规则，具备实际外联阻断能力。" "第2章" "持续核查出站规则，遵循最小化原则。"
    else
        add_result "2.13" "用户安全" "外联控制" "manual" "iptables OUTPUT 链未发现可确认的阻断规则（DROP/REJECT），或通过其他手段（firewalld rich rule、上网行为管理）控制外联，需人工核实。" "第2章" "通过防火墙/上网行为管理限制非必要的对外连接。"
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
    # 指导书3.5=日志读写控制/加密/变换/完整性校验，TM级数据存储加密。多层核查。
    local perm logrotate sudo_log integrity detail=""
    # 1. 读写控制：/var/log 权限
    perm="$(run_cmd "stat -c '%a' /var/log 2>/dev/null")"
    if [ -n "$perm" ] && [ "$perm" -le 750 ] 2>/dev/null; then
        detail="读写控制:/var/log权限$perm(≤750，合规); "
    elif [ -n "$perm" ]; then
        detail="读写控制:/var/log权限$perm(过宽，应≤750); "
    else
        detail="读写控制:无法获取/var/log权限; "
    fi
    # 2. 日志轮转(变换/完整性)：logrotate 是否配置 syslog/auth 日志轮转
    logrotate="$(run_cmd "grep -rlE 'syslog|auth|messages|/var/log' /etc/logrotate.d/ /etc/logrotate.conf 2>/dev/null | head -3")"
    if [ -n "$logrotate" ]; then
        detail="${detail}轮转:已配置logrotate($logrotate); "
    else
        detail="${detail}轮转:未发现syslog/auth日志轮转配置; "
    fi
    # 3. sudo 日志只读访问控制（visudo 审计用户只读）
    sudo_log="$(run_cmd "grep -E 'NOPASSWD.*(cat|grep|journalctl|less|more)' /etc/sudoers /etc/sudoers.d/* 2>/dev/null | head -2")"
    if [ -n "$sudo_log" ]; then
        detail="${detail}sudo只读:$sudo_log; "
    else
        detail="${detail}sudo只读:未配置审计用户日志只读sudo; "
    fi
    # 4. 完整性校验工具(AIDE/tripwire)
    if pkg_installed aide 2>/dev/null || pkg_installed tripwire 2>/dev/null; then
        integrity="已安装AIDE/tripwire"
        detail="${detail}完整性:$integrity"
    else
        detail="${detail}完整性:未安装AIDE/tripwire"
    fi
    # 综合判定
    if echo "$detail" | grep -q "读写控制:.*合规" && echo "$detail" | grep -q "轮转:已配置"; then
        add_result "3.5" "数据安全" "日志读写控制/轮转/完整性保护" "pass" "$detail" "第3章" "保持日志权限收紧、logrotate轮转与完整性校验。"
    elif echo "$detail" | grep -q "读写控制:.*过宽"; then
        add_result "3.5" "数据安全" "日志读写控制/轮转/完整性保护" "fail" "$detail 日志目录权限过宽，存在未授权读写风险。" "第3章" "收紧/var/log权限至750，配置logrotate轮转与完整性校验。"
    else
        add_result "3.5" "数据安全" "日志读写控制/轮转/完整性保护" "manual" "$detail 日志加密（LUKS/磁盘加密）与TM级数据存储加密需人工确认，脱敏变换配置需人工核查。" "第3章" "对日志采取读写控制、加密、变换、完整性校验，TM级数据存储加密。"
    fi
}

check_3_6_dlp() {
    # 指导书 3.6.3(3)：终端 DLP/管控客户端联动核查（边界设备侧仍需人工核实）
    local hit pkgs
    hit=""
    if command -v pgrep >/dev/null 2>&1; then
        hit="$(pgrep -a -i -f 'dlp|edr|endpoint|天御|终端安全' 2>/dev/null | head -3)"
    fi
    [ -z "$hit" ] && hit="$(ps aux 2>/dev/null | grep -iE 'dlp|edr|endpoint|天御|终端安全' | grep -v grep | head -3)"
    pkgs=""
    if command -v rpm >/dev/null 2>&1; then
        pkgs="$(rpm -qa 2>/dev/null | grep -iE 'dlp|edr|天御' | head -3 | tr '\n' ' ')"
    elif command -v dpkg >/dev/null 2>&1; then
        pkgs="$(dpkg -l 2>/dev/null | grep -iE 'dlp|edr|天御' | awk '{print $2}' | head -3 | tr '\n' ' ')"
    fi
    if [ -n "$hit" ] || [ -n "$pkgs" ]; then
        add_result "3.6" "数据安全" "边界防泄漏(DLP)终端联动" "pass" "检测到 DLP/终端管控客户端：${hit:-}${pkgs}。边界设备内容过滤与敏感内容识别策略仍需人工核实。" "第3章" "持续核查 DLP 覆盖主要外发通道并留存告警拦截日志。"
    else
        add_result "3.6" "数据安全" "边界防泄漏(DLP)终端联动" "fail" "本机未检测到 DLP/终端管控客户端进程或安装包；仅依赖本机防火墙/杀毒按指导书 3.6.2 不宜判定符合。" "第3章" "部署统一 DLP/终端管控客户端，并在边界设备启用内容过滤与敏感内容识别策略。"
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
    # 指导书3.8=按重要程度划分存储区块+设置访问权限。多层：独立分区+ACL访问权限。
    local mounts acl_info detail=""
    mounts="$(run_cmd "df -h 2>/dev/null | grep -E '/var$|/home$|/data$' | wc -l")"
    if [ "${mounts:-0}" -gt 0 ] 2>/dev/null; then
        detail="独立分区:/var /home /data 共${mounts}个; "
    else
        detail="独立分区:未检测到/var /home /data独立分区; "
    fi
    # 查关键数据目录ACL/权限
    local dir_perm
    dir_perm="$(run_cmd "stat -c '%a %n' /home /data /var/lib 2>/dev/null | head -3")"
    detail="${detail}目录权限:${dir_perm:-无法获取}"
    if [ "${mounts:-0}" -gt 0 ] 2>/dev/null; then
        add_result "3.8" "数据安全" "存储区块划分与访问权限" "manual" "$detail 已划分独立分区，但各存储区块ACL是否按用户/组分级授权需人工核查(getfacl)。" "第3章" "按重要程度划分存储区块并设置分级访问权限。"
    else
        add_result "3.8" "数据安全" "存储区块划分与访问权限" "manual" "$detail 建议将关键数据目录规划为独立分区并配置分级ACL。" "第3章" "按重要程度划分存储区块并设置分级访问权限。"
    fi
}

check_3_2_transferaudit() {
    local share_svc="" port="" audit_on=0
    svc_active smbd 2>/dev/null && share_svc="$share_svc smbd"
    svc_active smb 2>/dev/null && share_svc="$share_svc smb"
    svc_active vsftpd 2>/dev/null && share_svc="$share_svc vsftpd"
    port="$(run_cmd "ss -tlnp 2>/dev/null | grep -E ':445 |:21 '")"
    svc_active auditd && audit_on=1
    if [ -z "$share_svc" ] && [ -z "$port" ]; then
        add_result "3.2" "数据安全" "数据交互/文件传输统一管控和审计" "manual" "未检测到Samba/FTP等文件共享/传输服务，本机无此类数据交互通道；是否涉及其他应用内文件传输需人工确认。" "第3章" "数据交互与文件传输应统一管控并审计。"
    elif [ "$audit_on" -eq 1 ]; then
        add_result "3.2" "数据安全" "数据交互/文件传输统一管控和审计" "pass" "检测到文件共享/传输服务[${share_svc:-无} ${port:-}]，且auditd审计服务已启用，具备审计能力。" "第3章" "确保审计规则覆盖文件共享/传输行为并留存日志。"
    else
        add_result "3.2" "数据安全" "数据交互/文件传输统一管控和审计" "fail" "检测到文件共享/传输服务[${share_svc:-无} ${port:-}]，但auditd审计服务未启用，缺少统一审计。" "第3章" "启用auditd并配置文件传输/共享行为审计规则。"
    fi
}

check_3_3_wipe() {
    local tools=""
    command -v shred >/dev/null 2>&1 && tools="$tools shred"
    command -v wipe >/dev/null 2>&1 && tools="$tools wipe"
    if [ -n "$tools" ]; then
        add_result "3.3" "数据安全" "涉密载体降密级前写覆盖清除" "pass" "检测到数据写覆盖工具：$tools，具备写覆盖清除能力。" "第3章" "降密级前/演训结束后使用shred/wipe等工具对载体写覆盖清除。"
    else
        add_result "3.3" "数据安全" "涉密载体降密级前写覆盖清除" "fail" "未检测到shred/wipe等数据写覆盖工具，无法执行涉密载体写覆盖清除。" "第3章" "安装shred/wipe等工具，落实载体降密级前写覆盖清除。"
    fi
}

check_3_7_storagelogin() {
    local shadow_line="" locked=0
    shadow_line="$(run_cmd "grep -E '^(root|admin):' /etc/shadow 2>/dev/null")"
    if [ -n "$shadow_line" ] && echo "$shadow_line" | grep -qE '^[^:]+:[!*]'; then
        locked=1
    fi
    local audit_conf=""
    audit_conf="$(run_cmd "grep -E '^\s*(max_log_file|num_logs)' /etc/audit/auditd.conf 2>/dev/null | tr '\n' ' '")"
    local rot_cnt=""
    rot_cnt="$(run_cmd "grep -A5 -iE 'audit' /etc/logrotate.d/* /etc/logrotate.conf 2>/dev/null | grep -oE 'rotate[[:space:]]+[0-9]+' | grep -oE '[0-9]+' | head -1")"
    local retention_ok=0
    [ -n "$rot_cnt" ] && [ "$rot_cnt" -ge 26 ] 2>/dev/null && retention_ok=1
    if [ "$locked" -eq 1 ] && [ "$retention_ok" -eq 1 ]; then
        add_result "3.7" "数据安全" "数据存储管理登录增强/默认口令/审计/180天" "pass" "默认账户(root/admin)已锁定；审计留存rotate=$rot_cnt(约$((rot_cnt*7))天，≥180天)。auditd配置：${audit_conf:-无}。" "第3章" "保持默认账户锁定与审计日志≥180天留存。"
    elif [ "$locked" -eq 1 ]; then
        add_result "3.7" "数据安全" "数据存储管理登录增强/默认口令/审计/180天" "manual" "默认账户(root/admin)已锁定，但审计留存是否达180天未能确认(rotate=${rot_cnt:-未配置})，存储系统自身审计留存需人工核查。" "第3章" "确认存储系统审计日志留存≥180天。"
    else
        add_result "3.7" "数据安全" "数据存储管理登录增强/默认口令/审计/180天" "manual" "默认账户(root/admin)仍启用口令登录(${shadow_line:-无法读取shadow})，是否已改默认口令及管理登录增强措施(验证码等)需人工核查。" "第3章" "默认账户应改密或锁定，启用验证码等增强措施，审计留存≥180天。"
    fi
}

check_3_9_transencryption() {
    local tls
    tls="$(run_cmd "ss -tlnp 2>/dev/null | grep -E ':443 '")"
    if [ -n "$tls" ]; then
        add_result "3.9" "数据安全" "数据传输加密/路径/管控/日志/防泄漏" "pass" "检测到HTTPS/TLS加密监听(443)：$tls，数据传输采用加密通道。" "第3章" "保持数据传输使用HTTPS/TLS等加密协议。"
    else
        add_result "3.9" "数据安全" "数据传输加密/路径/管控/日志/防泄漏" "manual" "未检测到443端口HTTPS监听，本机数据传输是否加密(可能使用其他端口/协议或非本机传输)需人工确认。" "第3章" "数据传输应加密、路径合理、统一管控、留日志、具防泄漏措施。"
    fi
}

check_3_11_permgrading() {
    local shadow_perm
    shadow_perm="$(run_cmd "stat -c '%a' /etc/shadow 2>/dev/null")"
    if [ "$shadow_perm" = "600" ] || [ "$shadow_perm" = "640" ] || [ "$shadow_perm" = "400" ] || [ "$shadow_perm" = "440" ] || [ "$shadow_perm" = "0" ] || [ "$shadow_perm" = "000" ]; then
        add_result "3.11" "数据安全" "数据访问权限分级+审计" "pass" "/etc/shadow 权限为 $shadow_perm（已收紧，非644/666），关键数据访问权限设置合理。" "第3章" "关键数据文件保持最小权限(600/640)，并配置访问审计。"
    elif [ -n "$shadow_perm" ]; then
        add_result "3.11" "数据安全" "数据访问权限分级+审计" "fail" "/etc/shadow 权限为 $shadow_perm，过宽(应600/640)，存在未授权读取风险。" "第3章" "将/etc/shadow权限收紧为600或640。"
    else
        add_result "3.11" "数据安全" "数据访问权限分级+审计" "manual" "无法读取/etc/shadow权限，数据访问分级情况需人工核查。" "第3章" "关键数据文件应设最小权限并分级访问。"
    fi
}

check_3_13_dataclassification() {
    local luks=""
    luks="$(run_cmd "lsblk -o NAME,TYPE,FSTYPE,MOUNTPOINT 2>/dev/null | grep -iE 'crypt|LUKS'")"
    [ -z "$luks" ] && luks="$(run_cmd "dmsetup ls 2>/dev/null | grep -i crypt")"
    if [ -n "$luks" ]; then
        add_result "3.13" "数据安全" "数据各环节处理密级保密要求" "pass" "检测到磁盘加密(LUKS/crypt)映射：$luks，存储环节密级处理具备加密保护。" "第3章" "数据存储/处理/传输各环节保持密级相应加密保护。"
    else
        add_result "3.13" "数据安全" "数据各环节处理密级保密要求" "manual" "未检测到LUKS/crypt磁盘加密，数据传输加密(443)与访问控制等各环节是否满足密级要求需人工核查。" "第3章" "数据各环节处理应满足密级相应保密要求(存储/传输加密等)。"
    fi
}

# ============================================================
# 第4章 应用安全
# ============================================================

check_4_15_nla() {
    add_result "4.15" "应用安全" "文电文档签名验证/密级标识" "manual" "文电等文档专用业务处理系统的签名验证、密级标识功能属应用层能力，需人工核查系统是否具备数字签名验证与密级标识功能。指导书方法：检查gpg/证书签名工具、SELinux MLS(sestatus)与文件安全上下文(ls -Z)密级标记。" "第4章" "文电文档专用系统应具备签名验证与密级标识功能。"
}

check_4_16_roleseparation() {
    # 指导书4.16=远程管理应采取加密保护措施。原4.29的PermitRootLogin逻辑迁入此号。
    local rootlogin
    rootlogin="$(grep -Ei '^\s*PermitRootLogin' /etc/ssh/sshd_config 2>/dev/null | awk '{print $2}' | head -1)"
    local weak=""
    weak="$(grep -Ei '^\s*(Ciphers|MACs|KexAlgorithms)' /etc/ssh/sshd_config 2>/dev/null | grep -Ei '3des|arcfour|rc4|blowfish|cbc|sha1|diffie-hellman-group1')"
    local telnet
    telnet="$(run_cmd "ss -lntp 2>/dev/null | grep ':23 '")"
    if [ -n "$telnet" ]; then
        add_result "4.16" "应用安全" "远程管理加密保护" "fail" "检测到Telnet(23)明文管理服务在监听：$telnet。远程管理未全部采用加密保护。" "第4章" "停用Telnet，远程管理统一使用SSH等加密协议。"
    elif [ -n "$weak" ]; then
        add_result "4.16" "应用安全" "远程管理加密保护" "fail" "sshd_config 中配置了弱加密算法：$weak" "第4章" "禁用3DES/RC4/CBC/SHA1等弱算法，使用AES-GCM/ChaCha20等强算法。"
    elif [ "$rootlogin" = "no" ]; then
        add_result "4.16" "应用安全" "远程管理加密保护" "pass" "未检测到明文Telnet管理服务与弱加密算法，且 PermitRootLogin=no。" "第4章" "保持SSH加密管理，禁用Telnet，定期核查加密套件。"
    else
        add_result "4.16" "应用安全" "远程管理加密保护" "fail" "未检测到明文Telnet与弱加密算法，但 PermitRootLogin=${rootlogin:-未设置}，未明确禁止root直接远程登录。" "第4章" "在sshd_config设置 PermitRootLogin no。"
    fi
}

check_4_17_iprestrict() {
    # 指导书4.17=三权分立。原4.16的sudo逻辑迁入此号。
    local sudoers
    sudoers="$(run_cmd "wc -l < /etc/sudoers 2>/dev/null")"
    add_result "4.17" "应用安全" "三权分立" "manual" "sudo配置行数：${sudoers:-未知}，管理员/安全员/审计员的三权分立需人工核查账户与sudo授权划分（指导书方法：核查passwd/group中三员账户、:0:组、sudoers授权）。" "第4章" "划分系统管理、安全审计、业务操作等不同角色账户，避免权限过度集中。"
}

check_4_18_lockout() {
    # 指导书4.18=业务管理终端登录采取网络地址限制。原4.17的hosts.allow逻辑迁入此号。
    local allow sshd_allow
    allow="$(run_cmd "cat /etc/hosts.allow 2>/dev/null | grep -v '^#' | grep -v '^$'")"
    sshd_allow="$(grep -Ei '^\s*(AllowUsers|DenyUsers|AllowGroups|DenyGroups)' /etc/ssh/sshd_config 2>/dev/null)"
    if [ -n "$allow" ] || [ -n "$sshd_allow" ]; then
        add_result "4.18" "应用安全" "业务管理终端网络地址限制" "pass" "检测到管理访问来源限制：hosts.allow[${allow:-无}] sshd[${sshd_allow:-无}]。" "第4章" "持续维护允许管理访问的IP/账户白名单。"
    else
        add_result "4.18" "应用安全" "业务管理终端网络地址限制" "manual" "未在hosts.allow/sshd_config发现管理终端地址限制，请人工核实是否通过firewalld/安全组限制。" "第4章" "通过hosts.allow/firewalld/sshd AllowUsers限制管理访问来源IP。"
    fi
}

check_4_21_defaultport() {
    # 指导书4.21=基于专用物理部件/生物特征多因素数字证书认证。原SSH默认端口逻辑迁至4.22。
    add_result "4.21" "应用安全" "多因素数字证书认证" "manual" "基于专用物理部件或生物特征的多因素数字证书认证属应用层能力，需人工核查。指导书方法：检查pam_fprintd/pam_biometric模块、openssl ecparam(sm2)国密证书、证书库。" "第4章" "重要应用应支持多因素+数字证书认证。"
}

check_4_22_portseparation() {
    # 指导书4.22=应更改Web应用系统默认服务发布端口。原4.21的端口检查逻辑迁入并改为查Web默认端口。
    local webport
    webport="$(run_cmd "ss -lntp 2>/dev/null | grep -E ':80 |:443 |:8080 |:8443 '")"
    if [ -n "$webport" ]; then
        add_result "4.22" "应用安全" "更改Web默认服务端口" "fail" "Web服务仍使用默认端口(80/443/8080/8443)监听：$webport" "第4章" "将Web服务发布端口修改为非默认端口。"
    else
        add_result "4.22" "应用安全" "更改Web默认服务端口" "manual" "未在标准Web端口(80/443/8080/8443)检测到监听，可能未部署Web服务或已使用非默认端口，需人工确认。" "第4章" "Web服务应使用非默认端口发布。"
    fi
}
check_4_23_appdbsep() {
    # 指导书4.23=应分开设置管理端口与应用端口。原4.22逻辑迁入并改为管理/业务端口分离。
    add_result "4.23" "应用安全" "分开设置管理端口与应用端口" "manual" "管理端口与业务端口是否分离需人工核查。指导书方法：ss -tlnp检查管理端口(22/3389等)与业务端口(80/443等)是否隔离监听、firewalld区域划分。" "第4章" "管理端口与业务端口分离，管理端口仅限内网访问。"
}
check_4_24_appaudit() {
    # 指导书4.24=应用服务和数据存储应部署在不同服务器。原4.23逻辑迁入并改为应用/数据存储分离。
    add_result "4.24" "应用安全" "应用与数据存储分离部署" "manual" "应用服务与数据存储是否分离部署需人工核查。指导书方法：netstat查数据库端口ESTAB来源、mount查数据存储、hostname/ipconfig确认是否同机。" "第4章" "应用服务与数据存储建议部署于不同服务器。"
}

check_4_1_appbackup() {
    # 指导书v2.1 4.1.3 补充：应用层(备份配置) + 数据库层(binlog/mysqldump)。OS层备份详见4.32。
    local db_bak="" app_bak="" detail=""
    # 1. 数据库层：MySQL binlog + mysqldump定时任务
    if command -v mysql >/dev/null 2>&1; then
        local binlog
        binlog="$(run_cmd "mysql -e \"SHOW VARIABLES LIKE 'log_bin'\" 2>/dev/null | grep -i on")"
        [ -n "$binlog" ] && db_bak="MySQL binlog已开启(用于增量恢复)"
    fi
    local dump_cron
    dump_cron="$(run_cmd "grep -rlE 'mysqldump|dmap|sys_dump|pg_dump' /etc/cron.d /etc/cron.daily /var/spool/cron 2>/dev/null")"
    [ -n "$dump_cron" ] && db_bak="${db_bak:+$db_bak；}检测到数据库备份定时任务:$dump_cron"
    # 2. 应用层：应用配置中的备份参数
    app_bak="$(run_cmd "grep -rilE 'backup|备份|dump' /opt/*/conf/*.conf /opt/*/config/*.yml 2>/dev/null | head -3")"
    [ -n "$app_bak" ] && app_bak="应用配置含备份参数:$app_bak"
    detail="${db_bak:-数据库备份层未检测到binlog/备份定时任务}；${app_bak:-应用配置未检测到备份参数}；OS层备份详见4.32。"
    if [ -n "$db_bak" ] || [ -n "$app_bak" ]; then
        add_result "4.1" "应用安全" "应用系统软件应采取备份措施" "manual" "$detail 检测到备份相关配置，但备份有效性(是否近期/可恢复)需人工确认。" "第4章" "建立应用+数据库+OS三层备份机制并定期验证可恢复。"
    else
        add_result "4.1" "应用安全" "应用系统软件应采取备份措施" "manual" "$detail 应用自身备份功能/数据库备份策略需人工核查(含应用管理界面备份功能、mysqldump/RMAN/backupset)。" "第4章" "建立应用+数据库+OS三层备份机制并定期验证可恢复。"
    fi
}

check_4_6_waf() {
    # 指导书4.6=Web安全防护措施(8层通用方法)。OS层可查：WAF组件、安全响应头、Cookie属性。
    local waf="" sec_headers="" web_listening=""
    # 1. WAF组件
    if pkg_installed modsecurity 2>/dev/null || pkg_installed libmodsecurity3 2>/dev/null; then
        waf="modsecurity"
    fi
    if svc_active nginx 2>/dev/null && run_cmd "nginx -V 2>&1 | grep -qi modsecurity"; then
        waf="$waf nginx+modsecurity"
    fi
    # 2. 是否有Web服务在监听(80/443/8080)
    web_listening="$(run_cmd "ss -lntp 2>/dev/null | grep -E ':80 |:443 |:8080 '")"
    # 3. 安全响应头(curl本地Web服务)
    if [ -n "$web_listening" ] && command -v curl >/dev/null 2>&1; then
        local hdrs
        hdrs="$(run_cmd "curl -sI http://127.0.0.1 2>/dev/null")"
        if [ -n "$hdrs" ]; then
            local missing=""
            echo "$hdrs" | grep -qi 'X-Frame-Options' || missing="$missing X-Frame-Options"
            echo "$hdrs" | grep -qi 'X-Content-Type-Options' || missing="$missing X-Content-Type-Options"
            echo "$hdrs" | grep -qi 'Content-Security-Policy' || missing="$missing CSP"
            if [ -n "$missing" ]; then
                sec_headers="安全头缺失:$missing"
            else
                sec_headers="安全头齐全(X-Frame/X-Content-Type/CSP)"
            fi
        fi
    fi
    local detail="WAF:[${waf:-无}] Web监听:[${web_listening:+有}] $sec_headers"
    if [ -n "$waf" ]; then
        add_result "4.6" "应用安全" "Web应用安全防护措施" "manual" "$detail。WAF组件已部署，但输入校验/输出编码/SQL参数化/DB最小权限/错误信息控制等应用层8项需人工核查代码。" "第4章" "持续更新WAF规则，覆盖输入校验/输出编码/SQL注入/Cookie安全等。"
    elif [ -n "$web_listening" ]; then
        add_result "4.6" "应用安全" "Web应用安全防护措施" "manual" "$detail。部署了Web服务但未检测到WAF组件，安全头$([ -z \"\$sec_headers\" ] && echo '未获取' || echo '已查')。应用层防护需人工核查。" "第4章" "对外Web服务应部署WAF并配置安全响应头。"
    else
        add_result "4.6" "应用安全" "Web应用安全防护措施" "manual" "$detail。未检测到本机Web服务监听，可能未部署Web应用或采用前置设备，需人工核实。" "第4章" "对外Web服务应部署WAF防护。"
    fi
}

check_4_8_antitamper() {
    # 指导书v2.1 4.8.3 补充：专业防篡改系统 + 文件完整性 + Web目录权限。AIDE装了≠配置到位。
    local aide="" tamper_sys="" webdir="" detail=""
    # 1. 文件完整性工具
    if pkg_installed aide 2>/dev/null || pkg_installed tripwire 2>/dev/null; then
        aide="AIDE/tripwire已安装"
        # 查AIDE是否配置监控Web目录
        local aide_cfg
        aide_cfg="$(run_cmd "grep -E '/var/www|html' /etc/aide/aide.conf /etc/aide/aide.conf.d/* 2>/dev/null | head -2")"
        [ -n "$aide_cfg" ] && aide="$aide(已配置监控Web目录)" || aide="$aide(未配置监控Web目录)"
    fi
    # 2. 专业网页防篡改系统
    tamper_sys="$(run_cmd "systemctl list-units 2>/dev/null | grep -iE 'guard|tamper|protect|防篡改' | head -2")"
    [ -z "$tamper_sys" ] && tamper_sys="$(run_cmd "find /opt -maxdepth 2 -name '*guard*' -o -name '*tamper*' -o -name '*防篡*' 2>/dev/null | head -2")"
    [ -n "$tamper_sys" ] && tamper_sys="专业防篡改系统:$tamper_sys"
    # 3. Web目录权限/只读挂载
    if [ -d /var/www ]; then
        local mount_ro
        mount_ro="$(run_cmd "mount | grep '/var/www' | grep -i 'ro,'")"
        [ -n "$mount_ro" ] && webdir="/var/www只读挂载" || webdir="/var/www非只读挂载"
    fi
    detail="${aide:-无AIDE/tripwire}；${tamper_sys:-未检测到专业防篡改系统}；${webdir:-无Web目录}"
    if [ -n "$tamper_sys" ]; then
        add_result "4.8" "应用安全" "网页防篡改" "manual" "$detail 部署了专业防篡改系统，但Web目录权限与告警配置需人工确认。" "第4章" "保持防篡改系统运行，定期更新基线。"
    elif [ -n "$aide" ]; then
        add_result "4.8" "应用安全" "网页防篡改" "manual" "$detail AIDE已安装但需确认是否配置Web目录监控及专业防篡改系统。" "第4章" "部署专业防篡改系统或配置AIDE监控Web目录。"
    else
        add_result "4.8" "应用安全" "网页防篡改" "manual" "$detail 未检测到防篡改工具，请人工核实是否部署了网页防篡改系统。" "第4章" "对外Web服务应部署防篡改/文件完整性监控机制。"
    fi
}

check_4_12_sessionlimit() {
    # 指导书4.12=最大并发会话连接数限制。原会话超时逻辑迁至4.19。本号Top10补逻辑，先manual占位。
    add_result "4.12" "应用安全" "最大并发会话连接数限制" "manual" "应用最大并发会话连接数限制需人工核查。指导书方法：检查/proc/sys/net/core/somaxconn、file-max、ss并发连接数、sshd MaxSessions/MaxStartups。" "第4章" "配置最大并发会话与连接数限制，防止资源耗尽。"
}

check_4_29_remotemgmt() {
    # 指导书4.29=用户访问权限统一管理功能。原PermitRootLogin逻辑已迁至4.16。
    add_result "4.29" "应用安全" "用户访问权限统一管理" "manual" "用户访问权限统一管理需覆盖OS层+应用层统一认证平台。指导书v2.1 4.29.3补充：应用层(IAM/CAS/OAuth2.0/SSO: systemctl list-units | grep -iE \"iam\\|cas\\|oauth\\|sso\\|auth\"; grep -r \"sso\\|cas\\|oauth\\|auth.server\\|iam\\|统一认证\" /opt/*/conf/*.conf; 权限变更经统一平台审批可追溯)、OS层(sssd.conf/authconfig统一认证、sudoers授权)。" "第4章" "应用应接入统一认证/权限平台，权限变更集中受控。"
}

check_4_31_appbackup() {
    # 指导书4.31=最大并发会话/会话建立速率/单用户并发会话数。指导书方法：sshd MaxSessions/MaxStartups/MaxAuthTries、limits.conf maxlogins。
    local ssh_limit="" pam_limit="" detail=""
    # 1. sshd 并发会话限制
    ssh_limit="$(run_cmd "sshd -T 2>/dev/null | grep -iE 'maxsessions|maxstartups|maxauthtries'")"
    [ -z "$ssh_limit" ] && ssh_limit="$(grep -Ei '^\s*(MaxSessions|MaxStartups|MaxAuthTries)' /etc/ssh/sshd_config 2>/dev/null)"
    # 2. PAM/limits 单用户并发会话限制
    pam_limit="$(run_cmd "grep -rE 'maxlogins|maxsyslogins' /etc/security/limits.conf /etc/security/limits.d/ 2>/dev/null")"
    if [ -n "$ssh_limit" ] && [ -n "$pam_limit" ]; then
        detail="sshd:$ssh_limit; PAM limits:$pam_limit"
        add_result "4.31" "应用安全" "最大并发会话/速率/单用户并发限制" "pass" "$detail" "第4章" "保持并发会话与速率限制配置。"
    elif [ -n "$ssh_limit" ] || [ -n "$pam_limit" ]; then
        detail="sshd:[${ssh_limit:-未配置}]; PAM limits:[${pam_limit:-未配置}]"
        add_result "4.31" "应用安全" "最大并发会话/速率/单用户并发限制" "manual" "$detail 部分限制已配置，但应用层并发会话数/连接数限制需人工核查。" "第4章" "配置sshd MaxSessions/MaxStartups及limits.conf maxlogins。"
    else
        add_result "4.31" "应用安全" "最大并发会话/速率/单用户并发限制" "fail" "未检测到sshd MaxSessions/MaxStartups/MaxAuthTries及limits.conf maxlogins配置。" "第4章" "在sshd_config配置MaxSessions/MaxStartups/MaxAuthTries，limits.conf配置maxlogins。"
    fi
}

check_4_32_domesticsoftware() {
    # 指导书4.32=所有应用具备备份与恢复功能。原国产化逻辑迁至4.33。
    local bakdir bakfiles cron
    # 限定常见备份目录，排除 /tmp /var/cache /proc /sys 等缓存/系统目录，避免误判
    bakdir="$(run_cmd "find /backup /data /home /opt /var -maxdepth 3 -type d \\( -iname '*backup*' -o -iname '*bak' \\) 2>/dev/null | grep -vE '/tmp|/cache|/\\.cache' | head -3")"
    bakfiles="$(run_cmd "find /backup /data/backup /var/backup /home /opt -maxdepth 3 -type f \\( -name '*.bak' -o -name '*.sql.gz' -o -name '*backup*.tar.gz' \\) 2>/dev/null | head -5")"
    cron="$(run_cmd "grep -rilE 'mysqldump|dmap|sys_dump|backup' /etc/cron.d /etc/cron.daily /var/spool/cron 2>/dev/null")"
    if [ -n "$bakfiles" ]; then
        add_result "4.32" "应用安全" "应用备份与恢复功能" "manual" "在常见备份目录检测到备份文件：$bakfiles。请人工确认备份有效性（时间是否近期、是否可恢复）。" "第4章" "定期验证备份可恢复性，异地存储。"
    elif [ -n "$cron" ]; then
        add_result "4.32" "应用安全" "应用备份与恢复功能" "manual" "检测到与备份相关的定时任务：$cron。请人工确认备份执行情况与可恢复性。" "第4章" "定期验证备份可恢复性，异地存储。"
    elif [ -n "$bakdir" ]; then
        add_result "4.32" "应用安全" "应用备份与恢复功能" "manual" "检测到疑似备份目录：$bakdir，但未发现备份文件。请人工确认备份执行情况。" "第4章" "建立应用配置与数据的定期备份机制并验证可恢复。"
    else
        add_result "4.32" "应用安全" "应用备份与恢复功能" "manual" "未在常见位置检测到备份文件或定时备份任务，应用是否具备备份与恢复功能需人工核查（含应用自身备份功能、数据库备份、OS定时任务三层）。" "第4章" "建立应用配置与数据的定期备份机制并验证可恢复。"
    fi
}

check_4_33_domesticsoftware() {
    # 指导书4.33=应用软件基于国产自主可控软硬件自主开发。原4.32国产化逻辑迁入。
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
        add_result "4.33" "应用安全" "国产化软硬件自主开发" "pass" "操作系统本身即为国产操作系统：$KYLIN_TYPE；另检测到国产数据库/中间件相关软件包：$found。" "第4章" "持续核查关键应用软件与国产操作系统/CPU的适配兼容情况。"
    else
        add_result "4.33" "应用安全" "国产化软硬件自主开发" "manual" "操作系统本身即为国产操作系统：$KYLIN_TYPE；但未检测到已安装的国产数据库/中间件软件包，关键应用软件是否基于国产软硬件自主开发需人工核查。" "第4章" "核查关键应用软件与国产操作系统/CPU的适配兼容情况。"
    fi
}

check_4_25_logaudit() {
    # 指导书4.25=对所有访问行为和管理行为的日志审计功能，审计日志至少保留180天
    # 方法：systemctl auditd/rsyslog状态、/var/log/auth.log存在性、logrotate中audit的rotate/maxage≥180、ausearch/aureport可用性
    local svc_status="" log_file="" rotate_cfg="" tools=""
    # 1. 审计服务状态
    if svc_active auditd; then
        svc_status="auditd运行中"
        local rules
        rules="$(run_cmd 'auditctl -l 2>/dev/null | wc -l')"
        svc_status="$svc_status，已加载${rules:-0}条审计规则"
    elif svc_active rsyslog || svc_active syslog-ng; then
        svc_status="auditd未运行，但rsyslog/syslog-ng在运行"
    else
        svc_status="未检测到auditd或rsyslog/syslog-ng"
    fi
    # 2. 日志文件存在性
    for f in /var/log/audit/audit.log /var/log/auth.log /var/log/secure; do
        [ -f "$f" ] && log_file="$log_file $f" && break
    done
    log_file="${log_file:-未找到标准审计日志文件}"
    # 3. logrotate 180天留存
    local rotate180=""
    rotate180="$(run_cmd "grep -rE 'rotate|maxage' /etc/logrotate.d/ /etc/logrotate.conf 2>/dev/null | grep -iE 'audit|auth|secure|syslog' | head -3")"
    local rotate_days=""
    # 提取rotate次数估算天数（rotate×weekly≈天数，粗略）
    local rot_cnt
    rot_cnt="$(run_cmd "grep -A5 -iE 'audit|/var/log/secure|/var/log/auth.log' /etc/logrotate.d/* /etc/logrotate.conf 2>/dev/null | grep -E 'rotate[[:space:]]+[0-9]+' | grep -oE '[0-9]+' | head -1")"
    if [ -n "$rot_cnt" ] && [ "$rot_cnt" -ge 26 ] 2>/dev/null; then
        rotate_days="rotate=$rot_cnt(约$((rot_cnt*7))天，≥180天)"
    elif [ -n "$rot_cnt" ]; then
        rotate_days="rotate=$rot_cnt(约$((rot_cnt*7))天，不足180天)"
    fi
    # 4. 查询工具
    command -v ausearch >/dev/null 2>&1 && tools="$tools ausearch"
    command -v aureport >/dev/null 2>&1 && tools="$tools aureport"
    tools="${tools:-无ausearch/aureport}"
    # 综合判定
    local detail="$svc_status；日志文件:$log_file；轮转:$rotate_days${rotate180:+($rotate180)}；查询工具:$tools。"
    if echo "$svc_status" | grep -q "auditd运行中" && [ -n "$log_file" ] && echo "$rotate_days" | grep -q "≥180"; then
        add_result "4.25" "应用安全" "日志审计功能+180天留存" "pass" "$detail" "第4章" "保持审计服务运行与日志180天留存策略。"
    elif echo "$svc_status" | grep -q "未检测"; then
        add_result "4.25" "应用安全" "日志审计功能+180天留存" "fail" "$detail" "第4章" "安装并启用auditd，配置日志轮转留存不少于180天。"
    else
        add_result "4.25" "应用安全" "日志审计功能+180天留存" "manual" "$detail 日志留存是否达180天及审计规则覆盖面需人工确认。" "第4章" "确保auditd运行、审计规则覆盖访问/管理行为、日志留存≥180天。"
    fi
}

check_4_2_apppatch() {
    if [ "$PKG_MGR" = "apt" ]; then
        local cnt
        cnt="$(run_cmd "apt list --upgradable 2>/dev/null | grep -v '^Listing' | grep -c upgradable")"
        cnt="${cnt:-0}"
        if [ "$cnt" -gt 0 ] 2>/dev/null; then
            add_result "4.2" "应用安全" "应用系统补丁及时性" "fail" "检测到 $cnt 个可升级软件包(含应用)，补丁未安装到最新。" "第4章" "及时安装应用及系统安全补丁，升级包应经过安全性测试。"
        else
            add_result "4.2" "应用安全" "应用系统补丁及时性" "pass" "未检测到可升级软件包，应用补丁状态较新。" "第4章" "持续跟踪应用厂商安全公告并及时更新。"
        fi
    else
        local out
        out="$(run_cmd "yum check-update 2>/dev/null | grep -Ev '^$|^Loaded|^Last metadata|^Obsoleting|^Security:' | wc -l")"
        out="${out:-0}"
        if [ "$out" -gt 0 ] 2>/dev/null; then
            add_result "4.2" "应用安全" "应用系统补丁及时性" "fail" "yum check-update 检测到约 $out 个可更新软件包，应用补丁未安装到最新。" "第4章" "及时安装应用及系统安全补丁。"
        else
            add_result "4.2" "应用安全" "应用系统补丁及时性" "pass" "yum 未发现可更新软件包，应用补丁状态较新。" "第4章" "持续跟踪应用厂商安全公告并及时更新。"
        fi
    fi
}

check_4_3_trustedroot() {
    local tpm="" integ=""
    [ -d /sys/class/tpm/tpm0 ] && tpm="/sys/class/tpm/tpm0"
    if pkg_installed aide 2>/dev/null || pkg_installed tripwire 2>/dev/null; then
        integ="aide/tripwire"
    fi
    if [ -n "$tpm" ] || [ -n "$integ" ]; then
        add_result "4.3" "应用安全" "基于可信根可信验证" "pass" "检测到可信根/完整性监控：${tpm:+TPM}${integ:+ ${integ}}。" "第4章" "保持可信根/完整性监控运行，应用可信性破坏后报警。"
    else
        add_result "4.3" "应用安全" "基于可信根可信验证" "manual" "未检测到TPM或aide/tripwire完整性监控，应用基于可信根的可信验证能力需人工确认。" "第4章" "基于可信根对应用软件可信验证，破坏后报警。"
    fi
}

check_4_4_pubclssep() {
    local web="" db=""
    web="$(run_cmd "ss -tlnp 2>/dev/null | grep -E ':80 |:443 '")"
    db="$(run_cmd "ss -tlnp 2>/dev/null | grep -E ':3306 |:5432 |:6379 '")"
    if [ -n "$web" ] && [ -n "$db" ]; then
        add_result "4.4" "应用安全" "公共/涉密服务器分设、专用服务器专用" "fail" "同机同时运行Web服务[${web}]与数据库服务[${db}]，公共/涉密服务器未分设。" "第4章" "公共信息服务器与涉密/数据库服务器应分设部署。"
    elif [ -n "$web" ] || [ -n "$db" ]; then
        add_result "4.4" "应用安全" "公共/涉密服务器分设、专用服务器专用" "manual" "检测到Web服务[${web:-无}]或数据库服务[${db:-无}]单一类型，是否与其他服务器分设需人工结合网络架构确认。" "第4章" "公共/涉密服务器分设，专用服务器只提供专用服务。"
    else
        add_result "4.4" "应用安全" "公共/涉密服务器分设、专用服务器专用" "manual" "未检测到本机Web(80/443)或数据库(3306/5432/6379)服务监听，分设情况需人工结合网络架构确认。" "第4章" "公共/涉密服务器分设，专用服务器只提供专用服务。"
    fi
}

check_4_7_staticpublish() {
    local dyn="" webdir=0
    dyn="$(run_cmd "find /var/www /usr/share/nginx -type f \\( -name '*.php' -o -name '*.jsp' \\) 2>/dev/null | head -5")"
    [ -d /var/www ] && webdir=1
    [ -d /usr/share/nginx ] && webdir=1
    if [ -n "$dyn" ]; then
        add_result "4.7" "应用安全" "网站静态化发布" "fail" "检测到动态脚本文件：$dyn，网站未静态化发布。" "第4章" "网站宜以静态页面形式发布。"
    elif [ "$webdir" -eq 1 ]; then
        add_result "4.7" "应用安全" "网站静态化发布" "pass" "未检测到php/jsp动态脚本，网站以静态页面形式发布。" "第4章" "保持网站静态化发布。"
    else
        add_result "4.7" "应用安全" "网站静态化发布" "manual" "未检测到Web目录(/var/www、/usr/share/nginx)，是否部署网站及是否静态化发布需人工确认。" "第4章" "网站宜以静态页面形式发布。"
    fi
}

check_4_11_useraccess() {
    local sudo_group=""
    sudo_group="$(run_cmd "getent group sudo wheel 2>/dev/null | cut -d: -f1,4 | tr '\n' ' '")"
    add_result "4.11" "应用安全" "用户授权访问控制能力" "manual" "管理员组及成员[${sudo_group:-无sudo/wheel组}]。用户授权访问控制(RBAC)及数据库/应用层授权需人工核查。" "第4章" "应用应实现RBAC，数据库账号权限最小化。"
}

check_4_14_finegrained() {
    local selinux="" acl=""
    command -v getenforce >/dev/null 2>&1 && selinux="$(run_cmd getenforce)"
    acl="$(run_cmd "getfacl /etc/shadow 2>/dev/null | grep -E '^mask::|^user:[^:]+:|^group:[^:]+:'")"
    if [ "$selinux" = "Enforcing" ]; then
        add_result "4.14" "应用安全" "基于用户角色细粒度访问控制" "pass" "SELinux处于Enforcing模式，具备强制细粒度访问控制能力。" "第4章" "保持SELinux Enforcing并维护文件安全标签。"
    elif [ -n "$acl" ]; then
        add_result "4.14" "应用安全" "基于用户角色细粒度访问控制" "pass" "检测到文件扩展ACL配置：$acl，具备文件级细粒度访问控制。" "第4章" "保持文件ACL/角色权限配置。"
    else
        add_result "4.14" "应用安全" "基于用户角色细粒度访问控制" "manual" "未检测到SELinux Enforcing或文件扩展ACL，应用基于用户角色的细粒度访问控制(用户/进程/文件/表记录字段级)需人工核查。" "第4章" "应用应基于用户角色实现细粒度访问控制。"
    fi
}

check_4_19_loginfailure() {
    local pam_lock="" faillock_conf=""
    pam_lock="$(run_cmd "grep -E 'pam_faillock|pam_tally2' /etc/pam.d/system-auth /etc/pam.d/common-auth 2>/dev/null | grep -v '^\s*#'")"
    faillock_conf="$(run_cmd "grep -E '^\s*(deny|unlock_time)' /etc/security/faillock.conf 2>/dev/null")"
    if [ -n "$pam_lock" ] || [ -n "$faillock_conf" ]; then
        add_result "4.19" "应用安全" "登录失败处理(锁定/超时/自动退出)" "pass" "检测到登录失败锁定配置：${pam_lock:+PAM($pam_lock)}${faillock_conf:+ faillock.conf($faillock_conf)}。" "第4章" "保持登录失败锁定与自动退出配置。"
    else
        add_result "4.19" "应用安全" "登录失败处理(锁定/超时/自动退出)" "fail" "未检测到pam_faillock/pam_tally2登录失败锁定及faillock.conf deny/unlock_time配置。" "第4章" "在PAM中启用pam_faillock，配置登录失败锁定次数与解锁时间。"
    fi
}

check_4_30_unifiedmgmt() {
    local ssh_restrict="" rootlogin="" mgmt=""
    ssh_restrict="$(grep -Ei '^\s*(AllowUsers|DenyUsers)' /etc/ssh/sshd_config 2>/dev/null)"
    rootlogin="$(grep -Ei '^\s*PermitRootLogin' /etc/ssh/sshd_config 2>/dev/null | awk '{print $2}' | head -1)"
    [ "$rootlogin" = "no" ] && ssh_restrict="$ssh_restrict PermitRootLogin=no"
    mgmt="$(run_cmd "systemctl list-units --all 2>/dev/null | grep -iE 'cockpit|webmin|ansible' | head -3")"
    if [ -n "$ssh_restrict" ] || [ -n "$mgmt" ]; then
        add_result "4.30" "应用安全" "统一管理措施与远程管理限制" "pass" "检测到远程管理限制/集中管理：${ssh_restrict:+sshd($ssh_restrict)}${mgmt:+ 管理平台($mgmt)}。" "第4章" "保持统一管理措施并限制远程管理访问来源。"
    else
        add_result "4.30" "应用安全" "统一管理措施与远程管理限制" "fail" "未检测到sshd AllowUsers/DenyUsers/PermitRootLogin=no限制或cockpit/webmin/ansible等集中管理平台。" "第4章" "设置统一管理措施并限制远程管理(sshd AllowUsers/集中管理平台)。"
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
    # 2.3 已由 check_2_3_apppass 真实检测（空口令/自动登录）
    add_result "2.11_m" "用户安全" "口令认证硬件Key（双因子）" "na" "是否采用硬件Key等双因子认证方式需结合具体业务系统评估，本次系统级核查不适用。" "第2章" "对高权限账户建议启用硬件Key等双因子认证方式。"

    # ---- 第3章 补充（数据安全）----
    add_result "3.4" "数据安全" "涉密载体物理销毁(消磁/粉碎/溶解)" "manual" "确定销毁的涉密载体应采取消磁、粉碎、溶解、化浆、熔化等物理销毁，属管理流程层需人工核查销毁档案与影像。无OS层方法。" "第3章" "涉密载体销毁应双人操作、全程监控、留存销毁档案。"
    # 3.6 已由 check_3_6_dlp 真实检测（终端 DLP 客户端联动）
    add_result "3.12" "数据安全" "数据采集是否超出业务需求" "manual" "数据采集范围是否超出业务需求需人工核查。指导书方法：sc query查采集/同步服务、数据库表/字段与业务清单比对、采集日志抽查来源频率。" "第3章" "数据采集应限定在业务必要范围内。"
    add_result "3.14" "数据安全" "数据访问分级+审计+大数据统一管控" "manual" "大数据访问统一管控需覆盖OS层+大数据平台层+数据库层。指导书v2.1 3.14.3补充：大数据平台(HDFS ACL:hdfs dfs -getfacl; Hive权限:beeline -e \"SHOW GRANT\"; Apache Ranger策略页面; Kerberos:klist)、统一管控(ss -tlnp | grep ':50070\\|:10000\\|:8020'确认无直连旁路)、数据库审计覆盖大数据查询。" "第3章" "对大数据访问提供统一管控和访问控制。"

    # ---- 第4章 补充（应用安全，业务层，非OS层）----
    # 4.1 由 check_4_1_appbackup 函数查数据库备份层，此处不重复
    add_result "4.5" "应用安全" "公共信息服务器防DDoS能力" "manual" "防DDoS需多层协同。指导书v2.1 4.5.3补充：边界设备(登录防火墙核查DDoS清洗策略)、应用层限流(nginx limit_req/limit_conn、tomcat maxThreads/acceptCount、grep -r \"limit_req\\|limit_conn\" /etc/nginx/)、主机层(sysctl tcp_syncookies/tcp_max_syn_backlog、iptables limit)。" "第4章" "公共信息服务器应具备或依托DDoS防御能力。"
    add_result "4.9" "应用安全" "防范SQL注入/XSS攻击能力" "manual" "Web应用防SQL注入/XSS能力需人工核查。指导书方法：curl -I安全头(CSP/X-Frame)、curl构造SQLi观察返回。" "第4章" "Web应用应具备或依托WAF防SQL注入/XSS。"
    add_result "4.10" "应用安全" "执行代码有效性校验" "manual" "应用对上传/执行代码有效性校验需人工核查。指导书方法：nikto扫描、上传文件类型白名单配置。" "第4章" "应用应对可执行代码进行有效性与合法性校验。"
    add_result "4.13" "应用安全" "业务管理终端专设专用" "manual" "业务管理终端是否专设专用需现场人工核查。指导书方法：grep auth.log登录来源IP比对终端台账。" "第4章" "业务管理终端应指定专人专用并登记备案。"
    add_result "4.20" "应用安全" "自主设计协议/接口" "manual" "是否使用自主设计的网络服务/协议/接口需人工核查。指导书方法：find国密sm2/sm3/sm4配置、国产中间件/SDK。" "第4章" "关键业务通信宜使用自主设计协议/接口。"
    add_result "4.26" "应用安全" "账户异常行为监测" "manual" "账户异常登录/异常通信/异常文件上传监测需人工核查。指导书方法：核查应用日志/态势感知平台异常告警规则。" "第4章" "应用应具备账户异常行为监测能力。"
    add_result "4.27" "应用安全" "输入数据格式长度检查" "manual" "人机接口/网络通信/文件输入数据格式长度检查需人工核查。指导书方法：php upload_max_filesize/post_max_size、应用校验配置。" "第4章" "应用应对输入数据进行格式和长度检查。"
    add_result "4.28" "应用安全" "检测防御应用层攻击" "manual" "检测并防御SQL注入/网页篡改/XSS/DoS等应用层攻击需人工核查。指导书方法：curl安全头、curl构造SQLi/XSS观察返回。" "第4章" "应能检测并防御应用层攻击。"

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

json_esc() {
    local s="$1"
    s="${s//\\/\\\\}"
    s="${s//\"/\\\"}"
    s="${s//$'\n'/ }"
    s="${s//$'\r'/}"
    s="${s//$'\t'/ }"
    printf '%s' "$s"
}

report_rows_json() {
    local i first=1
    printf '['
    for ((i=1; i<=R_COUNT; i++)); do
        [ "$first" -eq 1 ] || printf ','
        first=0
        printf '{"ch":"%s","id":"%s","cat":"%s","title":"%s","status":"%s","detail":"%s","rec":"%s","guide":"%s"}' \
            "$(json_esc "${R_CHAPTER[$i]}")" "$(json_esc "${R_ID[$i]}")" "$(json_esc "${R_CAT[$i]}")" \
            "$(json_esc "${R_TITLE[$i]}")" "$(json_esc "${R_STATUS[$i]}")" "$(json_esc "${R_DETAIL[$i]}")" \
            "$(json_esc "${R_REC[$i]}")" "$(json_esc "${R_GUIDE[$i]}")"
    done
    printf ']'
}

REPORT_TAG=""
REPORT_TITLE="配置核查报告 - 中标麒麟/银河麒麟版"
report_meta_html() {
    echo "<div>系统：$(html_esc "$OS_NAME $OS_VER")　（$(html_esc "$KYLIN_TYPE")）</div>"
    echo "<div>主机名：$(html_esc "$HOSTNAME_STR")　内核：$(html_esc "$KERNEL_VER")</div>"
    echo "<div>核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>"
    echo "<div>参考标准：配置核查作业指导书v2.2</div>"
}

generate_html() {
    local outfile="$OUT_DIR/配置核查报告${REPORT_TAG:+_$REPORT_TAG}_${STAMP}.html"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in
            pass) pass=$((pass+1));;
            fail) fail=$((fail+1));;
            manual) manual=$((manual+1));;
            na) na=$((na+1));;
        esac
    done
    local total=$R_COUNT
    local ppass pfail pmanual pna
    ppass="$(awk -v a="$pass" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pfail="$(awk -v a="$fail" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pmanual="$(awk -v a="$manual" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    pna="$(awk -v a="$na" -v t="$total" 'BEGIN{if(t==0) printf "0.0"; else printf "%.1f", a*100/t}')"
    {
        cat <<HTMLHEAD
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${REPORT_TITLE}</title>
<style>
:root{
  --bg:#f5f7fa; --card:#fff; --ink:#1f2937; --muted:#6b7280; --line:#e5e7eb;
  --brand:#0f3057; --brand2:#1a3c6e;
  --pass:#15803d; --pass-bg:#ecfdf5; --pass-br:#bbf7d0;
  --fail:#b91c1c; --fail-bg:#fef2f2; --fail-br:#fecaca;
  --manual:#b45309; --manual-bg:#fffbeb; --manual-br:#fde68a;
  --na:#4b5563; --na-bg:#f3f4f6; --na-br:#e5e7eb;
}
*{box-sizing:border-box;}
body{margin:0;font:14px/1.65 "Segoe UI","Microsoft YaHei",system-ui,sans-serif;color:var(--ink);background:var(--bg);}
.wrap{max-width:1240px;margin:0 auto;padding:24px 20px 60px;}
header{background:linear-gradient(135deg,var(--brand) 0%,var(--brand2) 60%,#2563eb 130%);color:#fff;border-radius:12px;padding:26px 30px;margin-bottom:20px;}
header h1{margin:0 0 6px;font-size:22px;letter-spacing:.5px;}
header .sub{opacity:.85;font-size:13px;}
.meta{display:flex;flex-wrap:wrap;gap:8px 28px;margin-top:16px;padding-top:14px;border-top:1px solid rgba(255,255,255,.25);font-size:13px;}
.meta div{opacity:.95;}
.meta b{font-weight:600;opacity:.75;margin-right:6px;}
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
  <h1>${REPORT_TITLE}</h1>
  <div class="sub">参考标准：配置核查作业指导书v2.2</div>
  <div class="meta">
HTMLHEAD
        report_meta_html
        cat <<HTMLMID
  </div>
</header>

<div class="dash">
  <div class="stat s-pass" onclick="fset('pass')"><div class="num">$pass</div><div class="lbl">合规</div><div class="bar"><i style="width:$ppass%"></i></div><div class="pct">$ppass%</div></div>
  <div class="stat s-fail" onclick="fset('fail')"><div class="num">$fail</div><div class="lbl">不合规</div><div class="bar"><i style="width:$pfail%"></i></div><div class="pct">$pfail%</div></div>
  <div class="stat s-manual" onclick="fset('manual')"><div class="num">$manual</div><div class="lbl">需人工核查</div><div class="bar"><i style="width:$pmanual%"></i></div><div class="pct">$pmanual%</div></div>
  <div class="stat s-na" onclick="fset('na')"><div class="num">$na</div><div class="lbl">不适用</div><div class="bar"><i style="width:$pna%"></i></div><div class="pct">$pna%</div></div>
</div>

<div class="toolbar">
  <input id="q" type="text" placeholder="搜索编号 / 核查项 / 详情…">
  <div class="filters">
    <button class="fbtn on" data-f="all">全部<span class="count">$total</span></button>
    <button class="fbtn" data-f="fail">不合规<span class="count">$fail</span></button>
    <button class="fbtn" data-f="manual">需人工<span class="count">$manual</span></button>
    <button class="fbtn" data-f="pass">合规<span class="count">$pass</span></button>
    <button class="fbtn" data-f="na">不适用<span class="count">$na</span></button>
  </div>
</div>

<div class="panel">
<table id="tbl">
<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr></thead>
<tbody id="tb"></tbody>
</table>
<div class="empty" id="empty" style="display:none">没有匹配的核查项</div>
</div>

<footer>本报告由配置核查工具自动生成 · $(date '+%Y-%m-%d %H:%M:%S')</footer>
</div>

<script>
var DATA = 
HTMLMID
        report_rows_json
        cat <<HTMLFOOT
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
check_1_27_softwarelicensing
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
check_2_3_apppass
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
check_3_6_dlp
check_3_10_shares
check_3_8_storagepartition
check_3_2_transferaudit
check_3_3_wipe
check_3_7_storagelogin
check_3_9_transencryption
check_3_11_permgrading
check_3_13_dataclassification

# ---------- 第4章 应用安全 ----------
check_4_1_appbackup
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
check_4_33_domesticsoftware
check_4_25_logaudit
check_4_2_apppatch
check_4_3_trustedroot
check_4_4_pubclssep
check_4_7_staticpublish
check_4_11_useraccess
check_4_14_finegrained
check_4_19_loginfailure
check_4_30_unifiedmgmt

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
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_${STAMP}.xlsx" "配置核查报告 - 中标麒麟/银河麒麟版"

echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"
