#!/bin/bash
# ============================================================
# 麒麟（Linux）配置核查 - 一键运行全部（自动探测版）
#
# 到被测机上只跑这一个脚本：自动探测本机装了哪些数据库/中间件，
# 只核查探测到的组件，最后合并输出一份汇总 Excel(.xlsx)。
#
# 用法：sudo bash run_all.sh
#
# 连接参数可用环境变量覆盖（探测到但连接失败时按需提供）：
#   MYSQL_HOST=127.0.0.1 MYSQL_PORT=3306 MYSQL_USER=root MYSQL_PASS=xxx
#   REDIS_HOST=127.0.0.1 REDIS_PORT=6379 REDIS_PASS=xxx
#   DM_HOST=127.0.0.1 DM_PORT=5236 DM_USER=SYSDBA DM_PASS=xxx
#   MSSQL_HOST=127.0.0.1 MSSQL_PORT=1433 MSSQL_USER=sa MSSQL_PASS=xxx
#
# 探测误判时强制全跑：FORCE_ALL=1 bash run_all.sh
# 单独跳过某组件：SKIP_MYSQL=1 SKIP_REDIS=1 SKIP_DM=1 SKIP_MSSQL=1
#                 SKIP_NGINX=1 SKIP_TOMCAT=1
# ============================================================
set -u
cd "$(dirname "$0")"

# ---------- 探测辅助 ----------
port_listen() {  # port_listen 3306 -> 0=在监听
    local port="$1" out=""
    if command -v ss >/dev/null 2>&1; then
        out="$(ss -tln 2>/dev/null)"
    elif command -v netstat >/dev/null 2>&1; then
        out="$(netstat -tln 2>/dev/null)"
    fi
    [ -n "$out" ] && echo "$out" | grep -q ":${port} " && return 0
    return 1
}

proc_alive() {  # proc_alive mysqld -> 0=进程存在
    command -v pgrep >/dev/null 2>&1 || return 1
    pgrep -f "$1" >/dev/null 2>&1
}

dir_exists() {  # 任一目录存在即返回0
    local d
    for d in "$@"; do
        [ -d "$d" ] && return 0
    done
    return 1
}

detect_mysql() {
    command -v mysql >/dev/null 2>&1 && return 0
    command -v mysqld >/dev/null 2>&1 && return 0
    command -v mariadbd >/dev/null 2>&1 && return 0
    proc_alive 'mysqld|mariadbd' && return 0
    port_listen 3306 && return 0
    return 1
}

detect_redis() {
    command -v redis-server >/dev/null 2>&1 && return 0
    command -v redis-cli >/dev/null 2>&1 && return 0
    proc_alive redis-server && return 0
    port_listen 6379 && return 0
    return 1
}

detect_dm() {
    command -v disql >/dev/null 2>&1 && return 0
    proc_alive dmserver && return 0
    port_listen 5236 && return 0
    dir_exists /opt/dmdbms /dm8 /opt/dm8 && return 0
    return 1
}

detect_mssql() {
    proc_alive sqlservr && return 0
    port_listen 1433 && return 0
    dir_exists /opt/mssql && return 0
    command -v sqlservr >/dev/null 2>&1 && return 0
    return 1
}

detect_nginx() {
    command -v nginx >/dev/null 2>&1 && return 0
    proc_alive nginx && return 0
    [ -f /etc/nginx/nginx.conf ] && return 0
    return 1
}

detect_tomcat() {
    proc_alive 'catalina|java.*tomcat' && return 0
    dir_exists /usr/local/tomcat /opt/tomcat /usr/local/apache-tomcat* && return 0
    port_listen 8080 && return 0
    return 1
}

skip_flag() {  # skip_flag MYSQL -> SKIP_MYSQL=1 时返回0
    local name
    name="$(echo "$1" | tr '[:lower:]' '[:upper:]')"
    [ "${!name:-0}" = "1" ]
}

# ---------- 探测 ----------
echo "############################################################"
echo "#  麒麟配置核查 - 一键运行（自动探测本机组件）"
echo "############################################################"
echo ""
echo "==================== 组件探测结果 ===================="
[ "${FORCE_ALL:-0}" = "1" ] && echo "  （FORCE_ALL=1：跳过探测，全部核查）"

RUN_MYSQL=0; RUN_REDIS=0; RUN_DM=0; RUN_MSSQL=0; RUN_NGINX=0; RUN_TOMCAT=0
mark() { # mark 组件名 探测函数 环境开关名
    local label="$1" det="$2" runvar="$3" why=""
    if [ "${FORCE_ALL:-0}" = "1" ]; then
        why="强制"
    elif skip_flag "$4"; then
        echo "  [跳过] $label（SKIP 开关）"
        return
    elif "$det"; then
        :
    else
        echo "  [跳过] $label（未检测到：无命令/进程/端口/安装目录）"
        return
    fi
    eval "$runvar=1"
    echo "  [核查] $label${why:+（$why）}"
}

echo "  [核查] 麒麟操作系统（始终核查）"
mark "MySQL/MariaDB 数据库" detect_mysql     RUN_MYSQL   MYSQL
mark "Redis"                detect_redis     RUN_REDIS   REDIS
mark "达梦 DM8 数据库"      detect_dm        RUN_DM      DM
mark "SQL Server 数据库"    detect_mssql     RUN_MSSQL   MSSQL
mark "Nginx 中间件"         detect_nginx     RUN_NGINX   NGINX
mark "Tomcat 中间件"        detect_tomcat    RUN_TOMCAT  TOMCAT

# ---------- 执行 ----------
step=0
total=1
for v in RUN_MYSQL RUN_REDIS RUN_DM RUN_MSSQL RUN_NGINX RUN_TOMCAT; do
    [ "${!v}" = "1" ] && total=$((total + 1))
done

run_step() {  # run_step "标题" 脚本
    step=$((step + 1))
    echo ""
    echo "==================== ${step}/${total} $1 ===================="
    bash "$2" || echo "（$1 核查脚本执行异常，继续后续项）"
}

run_step "麒麟操作系统核查" check_kylin.sh
[ "$RUN_MYSQL"  = "1" ] && run_step "MySQL 数据库核查"    check_mysql.sh
[ "$RUN_REDIS"  = "1" ] && run_step "Redis 核查"          check_redis.sh
[ "$RUN_DM"     = "1" ] && run_step "达梦数据库核查"      check_dm.sh
[ "$RUN_MSSQL"  = "1" ] && run_step "SQL Server 核查"     check_sqlserver.sh
[ "$RUN_NGINX"  = "1" ] && run_step "Nginx 中间件核查"    check_nginx.sh
[ "$RUN_TOMCAT" = "1" ] && run_step "Tomcat 中间件核查"   check_tomcat.sh

# ---------- 汇总 ----------
echo ""
echo "==================== 合并生成汇总 Excel(.xlsx) ===================="
bash merge_xlsx.sh ./output

echo ""
echo "############################################################"
echo "#  核查完成。本轮实际核查 ${total} 项，报告："
ls -t output/*.html 2>/dev/null | head -"${total}" | sed 's/^/#    /'
echo "#  Excel 汇总报告位于项目根目录：配置核查汇总报告_*.xlsx"
echo "#"
echo "#  提示：若某组件已装但被跳过（连接参数特殊），可用："
echo "#    FORCE_ALL=1 sudo bash run_all.sh   # 强制全部核查"
echo "############################################################"
