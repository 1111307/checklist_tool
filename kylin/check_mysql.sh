#!/bin/bash
# ============================================================
# 配置核查工具 - MySQL/MariaDB 数据库版（无需Python，纯Bash实现）
# 参考标准：配置核查作业指导书正式版2026_4_1
#           配置核查表_v2.0.0.xlsx（数据库列：MySQL 共19项）
# 运行方式：sudo bash check_mysql.sh
# 连接方式：默认用 mysql 客户端 + ~/.my.cnf / /etc/my.cnf 的 [client] 段
#           也可用环境变量覆盖：
#             MYSQL_HOST=127.0.0.1 MYSQL_PORT=3306 \
#             MYSQL_USER=root MYSQL_PASS=xxx bash check_mysql.sh
# 输出文件：output/配置核查报告_MySQL_日期时间.html /.xls
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

# ---------- 连接参数 ----------
MYSQL_BIN="$(command -v mysql 2>/dev/null)"
# 读取 db_config.conf（KEY=VALUE 格式）
read_config() {
    local key="$1" conf="$SCRIPT_DIR/db_config.conf" val=""
    if [ -f "$conf" ]; then
        val="$(grep -E "^[[:space:]]*${key}[[:space:]]*=" "$conf" 2>/dev/null | tail -1 | sed 's/^[^=]*=[[:space:]]*//; s/[[:space:]]*$//')"
    fi
    printf '%s' "$val"
}

MYSQL_HOST="${MYSQL_HOST:-$(read_config MYSQL_HOST)}"
MYSQL_HOST="${MYSQL_HOST:-127.0.0.1}"
MYSQL_PORT="${MYSQL_PORT:-$(read_config MYSQL_PORT)}"
MYSQL_PORT="${MYSQL_PORT:-3306}"
MYSQL_USER="${MYSQL_USER:-$(read_config MYSQL_USER)}"
MYSQL_USER="${MYSQL_USER:-root}"
MYSQL_PASS="${MYSQL_PASS:-$(read_config MYSQL_PASS)}"

# ---------- 数据库版本/连接状态 ----------
DB_VERSION=""
CONN_OK=0     # 1=已连上  0=连不上
MYSQL_PRESENT=0  # 1=mysql客户端存在  0=不存在

# 执行一条查询，输出原始结果（-N -B 制表符分隔，无表头）
mysql_q() {
    local opts=(-h"$MYSQL_HOST" -P"$MYSQL_PORT" -u"$MYSQL_USER" --connect-timeout=5 -N -B)
    if [ -n "$MYSQL_PASS" ]; then
        "$MYSQL_BIN" "${opts[@]}" -p"$MYSQL_PASS" -e "$1" 2>/dev/null
    else
        "$MYSQL_BIN" "${opts[@]}" -e "$1" 2>/dev/null
    fi
}

# 取单个变量的值（SHOW VARIABLES LIKE ...）
mysql_var() {
    mysql_q "SHOW VARIABLES LIKE '$1';" 2>/dev/null | awk -F'\t' '{print $2}'
}

# 判断数据库版本是否已 EOL（停止维护），用于补丁项自动判定 fail
version_eol() {
    local v="$1" major mm mi
    major="$(echo "$v" | grep -oE '[0-9]+\.[0-9]+' | head -1)"
    mm="${major%%.*}"; mi="${major##*.}"
    if echo "$v" | grep -qi mariadb; then
        # MariaDB：10.6 及以下已 EOL，10.11 LTS / 11.x 仍在维护
        if [ "$mm" = "10" ] && [ "$mi" -le 6 ] 2>/dev/null; then echo "eol"; else echo "supported"; fi
    else
        # MySQL：5.x 已 EOL，8.0/8.4/9.x 仍在维护
        if [ "$mm" -le 5 ] 2>/dev/null; then echo "eol"; else echo "supported"; fi
    fi
}

# ---------- 指导书对照 ----------
guide_ref() {
    case "$1" in
        1.1)  echo "《配置核查作业指导书》第1章 系统安全 1.1：操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.7)  echo "《配置核查作业指导书》第1章 系统安全 1.7：数据库管理系统应删除冗余帐户，应设置不少于8个字符且字母大小写、数字及特殊字符混合编制的账户口令" ;;
        1.8)  echo "《配置核查作业指导书》第1章 系统安全 1.8：数据库管理系统应删除冗余存储过程" ;;
        1.9)  echo "《配置核查作业指导书》第1章 系统安全 1.9：数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" ;;
        1.10) echo "《配置核查作业指导书》第1章 系统安全 1.10：数据库管理系统应具有自主访问控制功能" ;;
        1.11) echo "《配置核查作业指导书》第1章 系统安全 1.11：数据库管理系统应具有备份和恢复功能" ;;
        1.12) echo "《配置核查作业指导书》第1章 系统安全 1.12：数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.13) echo "《配置核查作业指导书》第1章 系统安全 1.13：数据库管理系统的数据应和其它应用的数据分类独立存储" ;;
        1.15) echo "《配置核查作业指导书》第1章 系统安全 1.15：应具备数据库管理系统超级管理员远程登录限制能力" ;;
        1.16) echo "《配置核查作业指导书》第1章 系统安全 1.16：应具备数据库管理系统输入（参数）检查能力" ;;
        1.19) echo "《配置核查作业指导书》第1章 系统安全 1.19：应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.20) echo "《配置核查作业指导书》第1章 系统安全 1.20：数据库管理系统应配置安全策略" ;;
        1.21) echo "《配置核查作业指导书》第1章 系统安全 1.21：数据库管理系统应具有行级或列级审计功能" ;;
        1.22) echo "《配置核查作业指导书》第1章 系统安全 1.22：数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) echo "《配置核查作业指导书》第1章 系统安全 1.23：数据库管理系统仅为应用服务器提供访问服务" ;;
        1.24) echo "《配置核查作业指导书》第1章 系统安全 1.24：应具备日志审计能力，审计日志至少保留180天" ;;
        1.25) echo "《配置核查作业指导书》第1章 系统安全 1.25：检查是否具备边界保护能力，是否可以抗攻击、防篡改" ;;
        1.26) echo "《配置核查作业指导书》第1章 系统安全 1.26：检查是否有防病毒日志、补丁日志、记录相关信息，记录信息的完整、有效" ;;
        2.16) echo "《配置核查作业指导书》第2章 用户安全 2.16：检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" ;;
        *)    echo "《配置核查作业指导书》" ;;
    esac
}

add_result() {
    # $1 id  $2 category  $3 title  $4 status  $5 detail  $6 chapter  $7 recommendation
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

# 无法连接时的统一兜底：标记需人工核查；若 mysql 客户端都不存在则标记不适用
db_unreachable() {
    # $1 id $2 title $3 rec
    if [ "$MYSQL_PRESENT" = "0" ]; then
        add_result "$1" "系统安全-数据库" "$2" "na" "未检测到 mysql 客户端，无法对 MySQL 数据库进行自动核查，标记不适用。" "第1章" "$3"
    else
        add_result "$1" "系统安全-数据库" "$2" "manual" "无法连接 MySQL（${MYSQL_USER}@${MYSQL_HOST}:${MYSQL_PORT}），需人工登录核查。请检查 MYSQL_USER/MYSQL_PASS 环境变量或 ~/.my.cnf 配置。" "第1章" "$3"
    fi
}

html_esc() { local s="$1"; s="${s//&/&amp;}"; s="${s//</&lt;}"; s="${s//>/&gt;}"; printf '%s' "$s"; }
status_cn() { case "$1" in pass) echo "合规";; fail) echo "不合规";; manual) echo "需人工核查";; na) echo "不适用";; *) echo "$1";; esac; }
status_color() { case "$1" in pass) echo "#2e7d32";; fail) echo "#c62828";; manual) echo "#ef6c00";; na) echo "#757575";; *) echo "#000000";; esac; }

build_table_rows_by_status() {
    local want="$1" i
    for ((i=1; i<=R_COUNT; i++)); do
        [ "${R_STATUS[$i]}" != "$want" ] && continue
        local color scn; color="$(status_color "${R_STATUS[$i]}")"; scn="$(status_cn "${R_STATUS[$i]}")"
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

print_summary() {
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    echo "核查完成，共 $R_COUNT 项："
    echo "  合规(pass)：$pass    不合规(fail)：$fail    需人工核查(manual)：$manual    不适用(na)：$na"
}

generate_html() {
    local outfile="$OUT_DIR/配置核查报告_MySQL_${STAMP}.html"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    {
        cat <<HTMLHEAD
<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>MySQL 配置核查报告</title>
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
<h1>MySQL 数据库配置核查报告</h1>
<div class="meta">
<div>连接：$(html_esc "$MYSQL_USER@$MYSQL_HOST:$MYSQL_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")</div>
<div>主机名：$(hostname 2>/dev/null)　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>
<div>参考标准：配置核查作业指导书正式版2026_4_1 / 配置核查表_v2.0.0.xlsx</div>
</div>
<div class="summary">
<div class="card pass">合规<br>$pass</div>
<div class="card fail">不合规<br>$fail</div>
<div class="card manual">需人工核查<br>$manual</div>
<div class="card na">不适用<br>$na</div>
</div>
HTMLHEAD
        if [ "$fail" -gt 0 ]; then echo "<h2>一、未通过（$fail 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status fail; echo "</table>"; fi
        if [ "$manual" -gt 0 ]; then echo "<h2>二、需人工核查（$manual 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status manual; echo "</table>"; fi
        if [ "$na" -gt 0 ]; then echo "<h2>三、不适用（$na 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status na; echo "</table>"; fi
        if [ "$pass" -gt 0 ]; then echo "<h2>四、通过（$pass 项）</h2><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status pass; echo "</table>"; fi
        cat <<HTMLFOOT
</body>
</html>
HTMLFOOT
    } > "$outfile"
    echo "HTML报告已生成：$outfile"
}

generate_xls() {
    local outfile="$OUT_DIR/配置核查报告_MySQL_${STAMP}.xls"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    {
        cat <<XLSHEAD
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]>
<xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>
<x:Name>MySQL配置核查报告</x:Name>
<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
</x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
<![endif]-->
<style>
table{border-collapse:collapse;}
th,td{border:1px solid #999;padding:4px 6px;font-family:"Microsoft YaHei",Arial;font-size:12px;mso-number-format:"\@";}
th{background:#1a3c6e;color:#fff;font-weight:bold;}
</style>
</head>
<body>
<p>MySQL 数据库配置核查报告　连接：$(html_esc "$MYSQL_USER@$MYSQL_HOST:$MYSQL_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>
XLSHEAD
        if [ "$fail" -gt 0 ]; then echo "<p><b>一、未通过（$fail 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status fail; echo "</table>"; fi
        if [ "$manual" -gt 0 ]; then echo "<p><b>二、需人工核查（$manual 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status manual; echo "</table>"; fi
        if [ "$na" -gt 0 ]; then echo "<p><b>三、不适用（$na 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status na; echo "</table>"; fi
        if [ "$pass" -gt 0 ]; then echo "<p><b>四、通过（$pass 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>详情</th><th>建议</th><th>参考指导书</th></tr>"; build_table_rows_by_status pass; echo "</table>"; fi
        cat <<XLSFOOT
</body>
</html>
XLSFOOT
    } > "$outfile"
    echo "Excel(.xls)报告已生成：$outfile"
}

# ============================================================
# 数据库核查项（共19项，对应配置核查表 MySQL 列）
# ============================================================

check_1_1_patch() {
    local eol; eol="$(version_eol "$DB_VERSION")"
    if [ "$eol" = "eol" ]; then
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "fail" "当前版本 ${DB_VERSION} 已停止维护（EOL），无法再获得安全补丁。" "第1章" "升级到仍在维护的版本（MySQL 8.0+ / MariaDB 10.11 LTS），并及时安装补丁。"
    else
        add_result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "当前版本 ${DB_VERSION}（仍在维护期）。请人工核对是否为该大版本最新补丁版本（对比官方版本发布页）。" "第1章" "及时升级到最新稳定补丁版本，并留存升级记录。"
    fi
}

check_1_7_dbaccounts() {
    local anon empty weak
    anon="$(mysql_q "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='';")"
    # 空口令账户（排除 socket 认证账户）
    empty="$(mysql_q "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket');")"
    local vp_len vp_policy
    vp_len="$(mysql_var validate_password.length)"
    vp_policy="$(mysql_var validate_password.policy)"
    if [ -n "$anon" ]; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "存在匿名账户（无需口令即可登录）：$anon" "第1章" "删除所有匿名账户，禁止无口令访问。"
    elif [ -n "$empty" ]; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "存在空口令账户：$empty" "第1章" "为所有账户设置不少于8位、含大小写字母/数字/特殊字符的强口令。"
    elif [ -n "$vp_len" ] && [ "$vp_len" -lt 8 ] 2>/dev/null; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "口令策略 validate_password.length=${vp_len}（<8），不满足不少于8位要求。" "第1章" "设置 validate_password.length 不低于8，并启用强口令策略。"
    else
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "未发现匿名账户或空口令账户。口令策略：validate_password 长度=${vp_len:-未启用}，强度=${vp_policy:-未启用}。" "第1章" "持续定期核查账户，删除/锁定冗余与弱口令账户。"
    fi
}

check_1_8_dbprocedures() {
    local rout udf
    rout="$(mysql_q "SELECT COUNT(*) FROM information_schema.ROUTINES;")"
    udf="$(mysql_q "SELECT name FROM mysql.func;")"
    local risky
    risky="$(echo "$udf" | grep -Ei 'sys_exec|sys_eval|sys_bineval|sys_get' 2>/dev/null)"
    if [ -n "$risky" ]; then
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "fail" "存在高风险用户自定义函数(UDF)：$risky（可调用系统命令，等效 xp_cmdshell）。" "第1章" "删除非必要的高危 UDF（sys_exec/sys_eval 等），禁用 FILE 权限。"
    elif [ -n "$udf" ]; then
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "manual" "存在用户自定义函数(UDF)：$(echo "$udf" | tr '\n' ' ')，请人工确认是否均为业务必需。存储过程/函数总数：${rout:-未知}。" "第1章" "核查并删除冗余存储过程/UDF，遵循最小功能原则。"
    else
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "pass" "未发现用户自定义函数(UDF)；存储过程/函数总数：${rout:-0}。" "第1章" "持续控制存储过程与 UDF 数量，仅保留业务必需项。"
    fi
}

check_1_9_dbgrants() {
    # 持有危险全局权限的非系统账户：FILE/SUPER/PROCESS/SHUTDOWN/GRANT OPTION，或全局增删改查
    local dangerous
    dangerous="$(mysql_q "SELECT CONCAT(u.User,'@',u.Host) FROM mysql.user u WHERE u.User NOT IN ('root','mysql.sys','mysql.session','mysql.infoschema') AND (u.File_priv='Y' OR u.Super_priv='Y' OR u.Process_priv='Y' OR u.Shutdown_priv='Y' OR u.Grant_priv='Y' OR (u.Select_priv='Y' AND u.Insert_priv='Y' AND u.Update_priv='Y' AND u.Delete_priv='Y' AND u.Create_priv='Y' AND u.Drop_priv='Y'));")"
    local tab_priv col_priv db_priv
    db_priv="$(mysql_q "SELECT COUNT(*) FROM mysql.db;")"
    tab_priv="$(mysql_q "SELECT COUNT(*) FROM information_schema.TABLE_PRIVILEGES;")"
    col_priv="$(mysql_q "SELECT COUNT(*) FROM information_schema.COLUMN_PRIVILEGES;")"
    if [ -n "$dangerous" ]; then
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "fail" "存在持有危险全局权限（FILE/SUPER/PROCESS/GRANT 或全局增删改查）的非管理员账户：$dangerous" "第1章" "按最小权限原则收敛权限，回收危险全局权限，改用库/表/列级授权。"
    else
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "pass" "未发现持有危险全局权限的非管理员账户。已授权粒度：库级 ${db_priv:-0} 条 / 表级 ${tab_priv:-0} 条 / 列级 ${col_priv:-0} 条。" "第1章" "持续按最小权限原则，使用库级/表级/列级 GRANT 分配权限。"
    fi
}

check_1_10_dbaccess() {
    local bind skip
    bind="$(mysql_var bind_address)"
    skip="$(mysql_var skip_networking)"
    local listen
    # 只取本地监听地址列（ss 第4列），避免误匹配对端地址 0.0.0.0:*
    listen="$(ss -lntp 2>/dev/null | awk -v p="$MYSQL_PORT" '$4 ~ (":"p"$"){print $4}')"
    if [ "$skip" = "ON" ]; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "skip_networking=ON，数据库仅允许本机 socket 访问，具备自主访问控制基础。" "第1章" "保持网络监听关闭，仅本机访问。"
    elif echo "$listen" | grep -qE '^(0\.0\.0\.0|\*|\[::\])' 2>/dev/null; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "fail" "数据库端口监听在 0.0.0.0/::（对外暴露）：$listen。bind_address=${bind:-未设置}。" "第1章" "将 bind-address 改为内网地址或 127.0.0.1，配合防火墙限制访问来源。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "数据库未对外暴露监听（bind_address=${bind:-默认}），监听：${listen:-无}。" "第1章" "持续限制数据库监听地址与访问来源IP。"
    fi
}

check_1_11_dbbackup() {
    local logbin
    logbin="$(mysql_var log_bin)"
    local cron
    cron="$(grep -ril 'mysqldump\|mariabackup\|xtrabackup\|mysqlbackup' /etc/cron.d /etc/cron.daily /var/spool/cron 2>/dev/null)"
    if [ -n "$cron" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "检测到数据库备份计划任务：$cron" "第1章" "定期验证备份可恢复性，并异地存储备份。"
    elif [ "$logbin" = "ON" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "manual" "log_bin=ON（支持 PITR 恢复），但未在常见 crontab 位置发现自动全量备份任务，请人工确认备份策略。" "第1章" "建立定期全量备份（mysqldump/mariabackup），结合 binlog 实现可恢复。"
    else
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "fail" "未发现自动备份任务，且 log_bin=OFF，数据库无任何备份/恢复机制。" "第1章" "启用 binlog 并建立定期备份，定期演练恢复。"
    fi
}

check_1_12_dbauditplugin() {
    local glog glogfile audit
    glog="$(mysql_var general_log)"
    audit="$(mysql_q "SHOW VARIABLES LIKE 'audit%';" | tr '\n' ' ')"
    server_audit="$(mysql_q "SHOW VARIABLES LIKE 'server_audit%';" | tr '\n' ' ')"
    if [ -n "$audit" ] && ! echo "$audit" | grep -q 'audit_log_policy=OFF' 2>/dev/null; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "检测到审计插件配置：$audit" "第1章" "确保审计覆盖登录、权限变更、敏感数据操作等关键行为。"
    elif [ "$glog" = "ON" ]; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "general_log=ON（已开启通用查询日志，可记录 SQL 操作）。" "第1章" "建议升级为专用审计插件，审计粒度更细、开销更小。"
    else
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "未检测到审计插件（audit_log/server_audit），且 general_log=${glog:-OFF}。${server_audit:+检测到 server_audit 变量：$server_audit}" "第1章" "启用 MySQL 审计插件（audit_log 或 MariaDB server_audit）记录关键操作。"
    fi
}

check_1_13_dbisolation() {
    add_result "1.13" "系统安全-数据库" "数据库与业务隔离" "manual" "数据分类独立存储需结合业务架构人工核查：确认数据库是否与其它应用共享实例、敏感数据与非敏感数据是否分库/分表/分实例存储。" "第1章" "敏感数据与普通数据分类独立存储，必要时物理或网络隔离。"
}

check_1_15_dbremote() {
    local root_remote
    root_remote="$(mysql_q "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND LOCATE('%', Host) > 0;")"
    local bind skip fwrule
    bind="$(mysql_var bind_address)"
    skip="$(mysql_var skip_networking)"
    fwrule="$(iptables -S 2>/dev/null | grep -E "dport ${MYSQL_PORT}|${MYSQL_PORT}" | head -3)"
    if [ -n "$root_remote" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "fail" "root 账户允许从任意/网段主机登录（Host 含通配符 %）：$root_remote" "第1章" "禁止 root 远程登录，root 仅限 localhost；远程运维使用受限专用账户。"
    elif [ "$skip" = "ON" ] || [ "$bind" = "127.0.0.1" ] || [ -n "$fwrule" ]; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "pass" "已限制超管远程登录：skip_networking=${skip:-OFF}，bind_address=${bind:-默认}${fwrule:+，防火墙规则存在}。" "第1章" "持续保持 root 仅限本机/可信终端访问。"
    else
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "未发现 root@'%' 通配账户，但 bind_address=${bind:-默认} 且无针对性防火墙规则，请人工确认是否已限制超管远程登录来源。" "第1章" "通过防火墙/账户 Host 限制数据库远程访问来源。"
    fi
}

check_1_16_inputcheck() {
    add_result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "数据库输入（参数）检查属应用层能力（参数化查询/预编译语句），需人工或工具（如 sqlmap/AWVS）核查应用系统是否使用参数化查询防 SQL 注入。" "第1章" "应用层使用 PreparedStatement 参数化查询，避免拼接 SQL。"
}

check_1_19_dbdefaults() {
    local port
    port="$(mysql_var port)"
    local root_empty
    root_empty="$(mysql_q "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE User='root' AND (authentication_string='' OR authentication_string IS NULL) AND plugin NOT IN ('auth_socket','unix_socket','socket','auth_unix_socket');")"
    local issues=""
    [ "$port" = "3306" ] && issues="仍使用默认端口 3306"
    [ -n "$root_empty" ] && issues="${issues:+$issues；}root 账户为空口令"
    [ -z "$issues" ] && [ "$port" != "3306" ] && issues="未使用默认端口(端口=${port:-未知})"
    if [ -n "$root_empty" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$issues" "第1章" "修改默认端口、为 root 设置强口令或改用 socket 认证。"
    elif [ "$port" = "3306" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "数据库仍使用默认端口 3306 监听。" "第1章" "更换默认监听端口，禁用/加固 root 默认账户。"
    else
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/空口令问题（端口=${port:-未知}，root 无空口令）。" "第1章" "持续保持非默认端口与加固后的账户口令。"
    fi
}

check_1_20_dbpolicy() {
    local vp secure sqlmode
    vp="$(mysql_var validate_password.policy)"
    secure="$(mysql_var secure_file_priv)"
    sqlmode="$(mysql_var sql_mode)"
    local cnt=0
    [ -n "$vp" ] && cnt=$((cnt+1))
    [ -n "$secure" ] && cnt=$((cnt+1))
    if [ "$cnt" -ge 1 ]; then
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "pass" "检测到安全策略配置：口令策略=${vp:-未启用}；secure_file_priv=${secure:-未设置}；sql_mode=${sqlmode:-未设置}。" "第1章" "完善安全基线：启用 validate_password、限制 secure_file_priv、收紧 sql_mode。"
    else
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "fail" "未检测到明显安全策略（validate_password 未启用、secure_file_priv 未设置）。" "第1章" "配置数据库安全策略：口令复杂度、文件读写限制、SQL 模式收紧。"
    fi
}

check_1_21_dbaudit() {
    local audit glog
    audit="$(mysql_q "SHOW VARIABLES LIKE 'audit%';" | tr '\n' ' ')"
    glog="$(mysql_var general_log)"
    if [ -n "$audit" ]; then
        add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "检测到审计插件变量：$audit。行级/列级审计粒度需人工确认审计策略是否覆盖敏感行/列。" "第1章" "配置审计策略覆盖敏感表、行、列的关键操作。"
    elif [ "$glog" = "ON" ]; then
        add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "general_log=ON，可记录全部 SQL，但行级/列级细粒度审计能力需人工确认。" "第1章" "启用审计插件，配置行级/列级审计策略。"
    else
        add_result "1.21" "系统安全-数据库" "数据库操作审计" "fail" "未启用任何审计能力（无审计插件且 general_log=OFF），不具备行级/列级审计。" "第1章" "启用审计插件，配置行级/列级审计策略。"
    fi
}

check_1_22_dbauditretention() {
    add_result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立安全监控与审计措施（第三方审计平台/旁路审计）需人工核查部署情况。" "第1章" "部署独立审计监控，审计日志与业务数据分离存储。"
}

check_1_23_dbappsep() {
    local app_hosts
    app_hosts="$(mysql_q "SELECT CONCAT(User,'@',Host) FROM mysql.user WHERE Host='%';")"
    local bind skip
    bind="$(mysql_var bind_address)"
    skip="$(mysql_var skip_networking)"
    if [ -n "$app_hosts" ]; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "存在允许任意主机(%)访问的账户：$(echo "$app_hosts" | tr '\n' ' ')，请人工确认这些账户是否仅为应用服务器专用。" "第1章" "数据库账户 Host 精确限定为应用服务器IP，禁止使用 '%'。"
    elif [ "$skip" = "ON" ] || [ -n "$bind" ]; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "未发现 '%' 通配账户，且网络监听受限（skip_networking=${skip:-OFF}，bind_address=${bind:-默认}）。" "第1章" "持续确保数据库仅为应用服务器提供访问。"
    else
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "请人工确认数据库是否仅对应用服务器开放访问（无 '%' 通配账户，但需核实 bind-address 与防火墙）。" "第1章" "通过账户 Host 与网络策略限制数据库仅对应用服务器开放。"
    fi
}

check_1_24_logaudit180() {
    local expire_days expire_sec
    expire_days="$(mysql_var expire_logs_days)"
    expire_sec="$(mysql_var binlog_expire_logs_seconds)"
    local days=0
    if [ -n "$expire_sec" ] && [ "$expire_sec" -gt 0 ] 2>/dev/null; then
        days=$((expire_sec/86400))
    elif [ -n "$expire_days" ] && [ "$expire_days" -gt 0 ] 2>/dev/null; then
        days="$expire_days"
    fi
    # logrotate 对 mysql/mariadb 日志的轮转保留天数（rotate N ≈ 保留 N 份）
    local logrotate_cfg rot=0
    logrotate_cfg="$(grep -ril 'mysql\|mariadb' /etc/logrotate.d/ 2>/dev/null | head -1)"
    if [ -n "$logrotate_cfg" ]; then
        rot="$(grep -oE 'rotate[[:space:]]+[0-9]+' "$logrotate_cfg" 2>/dev/null | head -1 | grep -oE '[0-9]+')"
    fi
    if [ "$days" -ge 180 ] 2>/dev/null; then
        add_result "1.24" "系统安全-数据库" "日志审计留存180天" "pass" "binlog 保留时长约 ${days} 天（>=180天）：expire_logs_days=${expire_days:-无}，binlog_expire_logs_seconds=${expire_sec:-无}。" "第1章" "持续确保审计日志留存不少于180天。"
    elif [ -n "$rot" ] && [ "$rot" -ge 180 ] 2>/dev/null; then
        add_result "1.24" "系统安全-数据库" "日志审计留存180天" "pass" "检测到 logrotate 已配置 MySQL 日志轮转保留 ${rot} 天（>=180天）：$logrotate_cfg。" "第1章" "持续确保审计日志留存不少于180天。"
    elif [ "$days" -gt 0 ] 2>/dev/null; then
        add_result "1.24" "系统安全-数据库" "日志审计留存180天" "fail" "binlog 保留时长约 ${days} 天（<180天），且未检测到 logrotate 180天轮转（注意：MariaDB 10.3 的 expire_logs_days 上限为99）。" "第1章" "用 logrotate 轮转审计日志保留>=180天，或升级到支持 binlog_expire_logs_seconds 的版本。"
    else
        add_result "1.24" "系统安全-数据库" "日志审计留存180天" "manual" "未检测到明确的 binlog 过期配置，请人工核实审计日志留存时长是否满足180天。" "第1章" "配置 binlog 过期策略或 logrotate，保证审计日志留存>=180天。"
    fi
}

check_1_25_boundary() {
    local fw fwstate
    fw="$(iptables -S 2>/dev/null | head -20)"
    fwstate="$(firewall-cmd --state 2>/dev/null)"
    if [ -n "$fw" ] || [ "$fwstate" = "running" ]; then
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到本机防火墙（${fwstate:-iptables 规则存在}），但边界防护/WAF/网络隔离/防篡改需结合网络架构人工核查。" "第1章" "部署边界防护设备，限制数据库暴露面，防攻击、防篡改。"
    else
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到本机防火墙规则（iptables 为空、firewalld 未运行），数据库缺乏边界保护。" "第1章" "部署防火墙/WAF 等边界防护，限制数据库暴露面。"
    fi
}

check_1_26_avpatchlog() {
    local logerr general slow
    logerr="$(mysql_var log_error)"
    general="$(mysql_var general_log)"
    slow="$(mysql_var slow_query_log)"
    if [ -n "$logerr" ] && [ "$logerr" != "stderr" ] && { [ "$general" = "ON" ] || [ "$slow" = "ON" ]; }; then
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "manual" "错误日志已启用（$logerr），查询日志 general_log=${general:-OFF}、慢日志 slow_query_log=${slow:-OFF}。防病毒日志与补丁日志是否完整有效需人工核查。" "第1章" "确保错误日志、审计日志完整记录，补丁升级留有记录。"
    elif [ -n "$logerr" ] && [ "$logerr" != "stderr" ]; then
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "manual" "错误日志已启用（$logerr），但查询/审计日志未开启，日志记录不完整。" "第1章" "开启 general_log/slow_query_log 与审计日志，保留完整记录。"
    else
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "fail" "未配置错误日志到独立文件（log_error=${logerr:-未设置}），日志记录缺失。" "第1章" "配置 log_error 到独立文件，保留完整运行/补丁日志。"
    fi
}

check_2_16_patchlatest() {
    local eol; eol="$(version_eol "$DB_VERSION")"
    if [ "$eol" = "eol" ]; then
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "fail" "当前版本 ${DB_VERSION} 已停止维护（EOL），未完成补丁修复升级。" "第2章" "升级到仍在维护的版本，完成补丁修复并留存升级记录。"
    else
        add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "当前版本 ${DB_VERSION}（仍在维护期）。请人工核对是否已升级到最新版本并留存补丁升级记录。" "第2章" "定期升级到最新版本，记录补丁升级情况。"
    fi
}

# ============================================================
# 主流程
# ============================================================

echo "================================================================"
echo "  配置核查工具 - MySQL/MariaDB 数据库版（无需Python）"
echo "  参考标准：配置核查作业指导书正式版2026_4_1"
echo "================================================================"
echo "  连接：$MYSQL_USER@$MYSQL_HOST:$MYSQL_PORT"
echo ""

# 连接探测
if [ -z "$MYSQL_BIN" ]; then
    MYSQL_PRESENT=0
    DB_VERSION="未安装mysql客户端"
    echo "  未检测到 mysql 客户端，所有数据库项将标记为不适用。"
elif DB_VERSION="$(mysql_q "SELECT VERSION();" 2>/dev/null)"; [ -n "$DB_VERSION" ]; then
    MYSQL_PRESENT=1
    CONN_OK=1
    echo "  数据库版本：$DB_VERSION"
else
    MYSQL_PRESENT=1
    CONN_OK=0
    echo "  [!] 无法连接 MySQL（$MYSQL_USER@$MYSQL_HOST:$MYSQL_PORT），所有数据库项将标记为需人工核查。"
    echo "      请通过 MYSQL_USER/MYSQL_PASS 环境变量或 ~/.my.cnf 提供凭据。"
fi
echo ""
echo "开始检查，请稍候..."

# 逐项核查
if [ "$CONN_OK" = "1" ]; then
    check_1_1_patch
    check_1_7_dbaccounts
    check_1_8_dbprocedures
    check_1_9_dbgrants
    check_1_10_dbaccess
    check_1_11_dbbackup
    check_1_12_dbauditplugin
    check_1_13_dbisolation
    check_1_15_dbremote
    check_1_16_inputcheck
    check_1_19_dbdefaults
    check_1_20_dbpolicy
    check_1_21_dbaudit
    check_1_22_dbauditretention
    check_1_23_dbappsep
    check_1_24_logaudit180
    check_1_25_boundary
    check_1_26_avpatchlog
    check_2_16_patchlatest
else
    db_unreachable "1.1"  "数据库补丁程序"            "及时升级到最新稳定补丁版本。"
    db_unreachable "1.7"  "数据库账户管理"            "删除冗余账户，设置不少于8位强口令。"
    db_unreachable "1.8"  "数据库存储过程管理"        "删除冗余存储过程与高危UDF。"
    db_unreachable "1.9"  "数据库权限最小化"          "按最小权限原则分配账户权限。"
    db_unreachable "1.10" "数据库访问控制"            "限制数据库监听地址与访问来源IP。"
    db_unreachable "1.11" "数据库备份策略"            "建立定期备份并验证可恢复性。"
    db_unreachable "1.12" "数据库审计插件"            "启用数据库审计插件记录关键操作。"
    db_unreachable "1.13" "数据库与业务隔离"          "敏感数据与普通数据分类独立存储。"
    db_unreachable "1.15" "数据库远程访问控制"        "限制超管远程登录来源。"
    db_unreachable "1.16" "输入验证/SQL注入防护"      "应用层使用参数化查询防SQL注入。"
    db_unreachable "1.19" "数据库默认账户/端口"       "修改默认端口，禁用/加固默认账户。"
    db_unreachable "1.20" "数据库安全策略"            "配置口令复杂度与文件读写限制。"
    db_unreachable "1.21" "数据库操作审计"            "开启行级/列级审计。"
    db_unreachable "1.22" "数据库审计日志留存"        "独立审计监控，日志分离存储。"
    db_unreachable "1.23" "应用与数据库账户分离"      "数据库仅为应用服务器提供访问。"
    db_unreachable "1.24" "日志审计留存180天"         "审计日志留存不少于180天。"
    db_unreachable "1.25" "边界保护抗攻击防篡改"      "部署边界防护，限制暴露面。"
    db_unreachable "1.26" "防病毒/补丁日志记录"       "保留完整运行/补丁日志。"
    db_unreachable "2.16" "补丁修复升级到最新"        "定期升级到最新版本并留存记录。"
fi

# 生成报告
echo ""
echo "================================================================"
print_summary
generate_html
generate_xls
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_MySQL_${STAMP}.xlsx" "MySQL 数据库配置核查报告"
echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"
