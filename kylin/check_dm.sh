#!/bin/bash
# ============================================================
# 配置核查工具 - 达梦数据库(DM8) 版（无需Python，纯Bash实现）
# 参考标准：配置核查作业指导书v2.0.0
#           配置核查表_v2.0.0.xlsx（数据库列：达梦 共19项）
# 运行方式：sudo bash check_dm.sh
# 连接方式：默认 disql 连接 SYSDBA/SYSDBA@127.0.0.1:5236；可用环境变量覆盖：
#             DM_HOST=127.0.0.1 DM_PORT=5236 DM_USER=SYSDBA DM_PASS=xxx bash check_dm.sh
# 输出文件：output/配置核查报告_达梦_日期时间.html /.xls
# !!! 注意：本脚本尚未经真实达梦实例实测，disql 输出解析按 DM8 语法编写，
#     SQL 视图/参数名（V$VERSION / V$DM_INI / DBA_USERS 等）需在实际环境核对。
# ============================================================
set -u

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUT_DIR="$SCRIPT_DIR/output"
mkdir -p "$OUT_DIR"
STAMP="$(date +%Y%m%d_%H%M%S)"

# 零依赖 .xlsx 生成器（可选，需要 zip 命令；缺失时自动降级为 .xls）
[ -f "$SCRIPT_DIR/lib_xlsx.sh" ] && source "$SCRIPT_DIR/lib_xlsx.sh"

R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=(); R_METHOD=()
R_COUNT=0

# disql 客户端定位（Docker 场景通常在 /opt/dmdbms/bin/disql，且需 LD_LIBRARY_PATH）
DISQL="$(command -v disql 2>/dev/null)"
if [ -z "$DISQL" ]; then
    for p in /opt/dmdbms/bin/disql /usr/local/dmdbms/bin/disql; do
        [ -x "$p" ] && DISQL="$p" && break
    done
fi
# 读取 db_config.conf（KEY=VALUE 格式）
read_config() {
    local key="$1" conf="$SCRIPT_DIR/db_config.conf" val=""
    if [ -f "$conf" ]; then
        val="$(grep -E "^[[:space:]]*${key}[[:space:]]*=" "$conf" 2>/dev/null | tail -1 | sed 's/^[^=]*=[[:space:]]*//; s/[[:space:]]*$//')"
    fi
    printf '%s' "$val"
}

DM_HOST="${DM_HOST:-$(read_config DM_HOST)}"
DM_HOST="${DM_HOST:-127.0.0.1}"
DM_PORT="${DM_PORT:-$(read_config DM_PORT)}"
DM_PORT="${DM_PORT:-5236}"
DM_USER="${DM_USER:-$(read_config DM_USER)}"
DM_USER="${DM_USER:-SYSDBA}"
DM_PASS="${DM_PASS:-$(read_config DM_PASS)}"

# disql 依赖 libdisql_dll.so，若在同一目录则加入 LD_LIBRARY_PATH
if [ -n "$DISQL" ]; then
    _dmdir="$(dirname "$DISQL")"
    [ -f "$_dmdir/libdisql_dll.so" ] && export LD_LIBRARY_PATH="$_dmdir${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
fi

DB_VERSION=""
CONN_OK=0
DM_PRESENT=0

# 执行 SQL，返回 disql 原始输出（-S 静默模式，从 stdin 读 SQL）
dm_q() {
    printf '%s\n' "$1" | "$DISQL" -S "$DM_USER"/"$DM_PASS"@"$DM_HOST":"$DM_PORT" 2>/dev/null
}

# 取 V$DM_INI 参数值（disql 输出为带表头表格，取数据行末列）
dm_ini() {
    dm_q "SELECT PARA_VALUE FROM V\$DM_INI WHERE PARA_NAME='$1';" 2>/dev/null | awk 'NF>=2 && $1 ~ /^[0-9]+$/ {print $NF; exit}'
}

# 取 COUNT 值
dm_count() {
    dm_q "$1" 2>/dev/null | awk 'NF>=2 && $1 ~ /^[0-9]+$/ {print $NF; exit}'
}

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

method_of() {
    # 评审整改第4条：验证过程/方法（源自指导书 v2.0.0 逐条提取，_gen_method_map.py 生成）
    case "$1" in
        1.1) printf %s "1.1.1 核查操作系统安装补丁情况：图形界面/文档核查；1.1.2 核查数据库补丁情况：图形界面/文档核查；1.1.3 核查中间件补丁情况：图形界面/文档核查" ;;
        1.2) printf %s "1.2.1 Windows7、windowsXP：wmic /namespace:\\\\root\\\\SecurityCenter2 path ；1.2.2 中标麒麟、银河麒麟：getstatus、setstatus -p disable、setstatus enable softmode" ;;
        1.3) printf %s "1.3.1 Windows7、windowsXP：services.msc；1.3.2 中标麒麟、银河麒麟：图形界面/文档核查" ;;
        1.4) printf %s "1.4.1 Windows7、WindowsXP：Windows 防火墙、netsh advfirewall show allprofiles；1.4.2 中标麒麟、银河麒麟：systemctl status firewalld、firewall-cmd --state、firewall-cmd --list-all" ;;
        1.5) printf %s "1.5.1 Windows7、WindowsXP：QOS数据包计划程序；1.5.2 中标麒麟、银河麒麟：ip link show、nmcli connection show、systemctl status bluetooth" ;;
        1.6) printf %s "1.6.1 Windows7、WindowsXP：启用或关闭Windows功能、Telnet客户端、Telnet服务器；1.6.2 中标麒麟、银河麒麟：图形界面/文档核查" ;;
        1.7) printf %s "1.7.1 Mysql：mysql -uroot -p、SELECT User, Host FROM mysql.user;、INSTALL PLUGIN validate_password SONAME 'vali" ;;
        1.8) printf %s "1.8.1 Mysql：mysql -uroot -p、DROP PROCEDURE IF EXISTS ‘old_backup_procedur、vi ~/.bashrc；1.8.2 SQLServer：USE YourDatabaseName;、sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.9) printf %s "1.9.1 Mysql：mysql -uroot -p、SHOW GRANTS FOR 'username'@'localhost';、vi ~/.bashrc；1.9.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.10) printf %s "1.10.1 Mysql：mysql -uroot -p、SELECT user, host FROM mysql.user;、SHOW GRANTS FOR 'user'@'host';；1.10.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.11) printf %s "1.11.1 Mysql：mysql -uroot -p、mysqladmin -uroot -p create restore_verify、mysql restore_verify < 备份文件路径\\dump.sql" ;;
        1.12) printf %s "1.12.1 Mysql：mysql -uroot -p、SET SESSION sql_safe_updates = 1;、SELECT @@sql_safe_updates;；1.12.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.13) printf %s "1.13.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE  ‘port’;、SHOW VARIABLES LIKE ‘datadir’;；1.13.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.14) printf %s "1.14.1 Nginx：图形界面/文档核查；1.14.2 Tomcat：<Server port=\"8527\" shutdown=\"DangerousShutdo" ;;
        1.15) printf %s "1.15.1 MySQL：SELECT user, host FROM mysql.user WHERE user 、DROP USER 'root'@'%';、vi ~/.bashrc；1.15.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.16) printf %s "1.16.1 MySQL：mysql -uroot -p、vi ~/.bashrc、vi ~/.bash_profile；1.16.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'；1.16.3 达梦：disql SYSDBA/SYSDBA@localhost:5236、su  -用户名、cd  /达梦数据库安装目录下的 bin 目录" ;;
        1.17) printf %s "1.17.1 Windows操作系统（Win7/XP）：services.msc、sc stop 服务名、sc config 服务名 start= disabled；1.17.2 中标麒麟、银河麒麟：firewall-cmd --list-all、ufw status verbose、systemctl list-units --type=service --state=r" ;;
        1.18) printf %s "1.18.1 Windows7、WindowsXP：fsutil quota query C:；1.18.2 中标麒麟、银河麒麟：cat /etc/cgconfig.conf、systemd-cgtop、cat /etc/security/limits.conf" ;;
        1.19) printf %s "1.19.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE 'port';、RENAME USER 'root'@'localhost' TO 'new_admin'；1.19.2 SQLServer：ALTER LOGIN sa WITH PASSWORD = 'YourNewPasswo、sqlcmd -S loca" ;;
        1.20) printf %s "1.20.1 MySQL：mysql -uroot -p、INSTALL PLUGIN validate_password SONAME 'vali、UNINSTALL PLUGIN  validate_password;；1.20.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'、sp_configure 'user" ;;
        1.21) printf %s "1.21.1 MySQL：mysql -uroot -p、SET GLOBAL general_log = 'ON'、SHOW VARIABLES LIKE '%general_log%';；1.21.2 SQLServer：sqlcmd -S localhost -U SA -P 'SA账户密码'" ;;
        1.22) printf %s "1.22.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE '%general_log%';、SET GLOBAL general_log = 'ON';；1.22.2 SQLServer：图形界面/文档核查" ;;
        1.23) printf %s "1.23.1 MySQL：mysql -uroot -p、SHOW VARIABLES LIKE 'bind_address';、SHOW GRANTS FOR 'app_user'@'应用服务器IP';；1.23.2 SQLServer：图形界面/文档核查" ;;
        1.24) printf %s "1.24.1 操作系统日志审计核查：图形界面/文档核查；1.24.2 数据库日志审计核查：图形界面/文档核查" ;;
        1.25) printf %s "1.25.1 操作系统抗攻击与防篡改核查：图形界面/文档核查；1.25.2 数据库抗攻击与防篡改核查：图形界面/文档核查" ;;
        1.26) printf %s "1.26.1 操作系统安全日志核查：图形界面/文档核查；1.26.2 数据库安全日志核查：图形界面/文档核查" ;;
        1.27) printf %s "1.27.1 操作系统：图形界面/文档核查；1.27.2 数据库管理系统：图形界面/文档核查；1.27.3 中间件：图形界面/文档核查；1.27.4 办公软件：图形界面/文档核查" ;;
        2.1) printf %s "2.1.1 服务器登录口令核查：图形界面/文档核查；2.1.2 用户计算机登录口令核查：图形界面/文档核查" ;;
        2.2) printf %s "2.2.1 Windows7/WindowXP：systeminfo | findstr /B /C:\"[补丁程序]\" /C:\"[更新]\"、systeminfo、wmic qfe list brief /format:table" ;;
        2.3) printf %s "2.3.1 Windows7/WindowXP：net user、net user 用户名、HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows；2.3.2 中标麒麟/银河麒麟：passwd、sudo passwd 目标用户名" ;;
        2.4) printf %s "2.4.1 Windows7/WindowXP：cmd、net user、net user 用户名；2.4.2 中标麒麟/银河麒麟：cat /etc/passwd、sudo awk -F: '{print $1}' /etc/shadow、userdel -r 重复用户名；2.4.3 核查口令互不相同：cat /etc/shadow、sudo passwd 用户名" ;;
        2.5) printf %s "2.5.1 Windows7/WindowXP：netstat -ano；2.5.2 中标麒麟/银河麒麟：sudo systemctl list-unit-files --type=service、sudo systemctl is-enabled <服务名>、sudo systemctl is-active <服务名>" ;;
        2.6) printf %s "2.6.1 Windows7/WindowXP：netsh advfirewall show allprofiles、wf.msc、gpedit.msc；2.6.2 中标麒麟/银河麒麟：sudo kylin-firewall -s on、sudo kylin-firewall -g、sudo kylin-firewall -A -n Deny-Telnet -p tcp " ;;
        2.7) printf %s "2.7.1 Windows7/WindowXP：图形界面/文档核查；2.7.2 中标麒麟/银河麒麟：md5sum /bin/bash、sudo apt-get install auditd、sudo auditctl -w /etc/passwd -p wa" ;;
        2.8) printf %s "2.8.1 Windows7/WindowXP：图形界面/文档核查；2.8.2 中标麒麟/银河麒麟：nmcli device status、sudo nmcli device set wlan0 managed no、sudo nmcli connection delete 无线连接名" ;;
        2.9) printf %s "2.9.1 Windows7/WindowXP：appwiz.cpl、发行者（Publisher）；2.9.2 中标麒麟/银河麒麟：图形界面/文档核查" ;;
        2.10) printf %s "2.10.1 Windows7/WindowXP：gpedit.msc、devmgmt.msc、eventvwr.msc；2.10.2 中标麒麟/银河麒麟：lsusb" ;;
        2.11) printf %s "2.11.1 Windows7/WindowXP：net accounts、HKEY_CURRENT_USER\\Control Panel\\Desktop；2.11.2 中标麒麟/银河麒麟：gsettings get org.gnome.desktop.session idle-" ;;
        2.12) printf %s "2.12.1 Windows7/WindowXP：devmgmt.msc、网络适配器（Network adapters）、netsh wlan show interfaces；2.12.2 中标麒麟/银河麒麟：图形界面/文档核查" ;;
        2.13) printf %s "2.13.1 Windows7/WindowXP：ipconfig/all；2.13.2 中标麒麟/银河麒麟：sudo netstat -tunap" ;;
        2.14) printf %s "2.14.1 Windows7/WindowXP：eventvwr.msc；2.14.2 中标麒麟/银河麒麟：sudo systemctl status auditd --no-pager" ;;
        2.15) printf %s "2.15.1 Windows7/WindowXP：计算机配置 → Windows 设置 → 安全设置 → 账户策略/本地策略、auditpol /get /category:*、net accounts；2.15.2 中标麒麟/银河麒麟：systemctl status firewalld、ufw status verbose、firewall-cmd --list-all" ;;
        2.16) printf %s "2.16.1 核查操作系统安装补丁情况：图形界面/文档核查；2.16.2 核查数据库补丁修复和升级到最新版本：图形界面/文档核查；2.16.3 应用软件补丁修复和升级到最新版本：图形界面/文档核查" ;;
        3.1) printf %s "3.1.1 Win7\\WinXP：manage-bde -status、cipher /u /n；3.1.2 中标麒麟、银河麒麟：sudo apt-get install cryptsetup、sudo cryptsetup luksOpen /dev/sdb1 crypt_data；3.1.3 通用核查方法（数据库层与应用层）：图形界面/文档核查" ;;
        3.2) printf %s "3.2.1 Win7\\WinXP：auditpol /get /category:*、auditpol /get /subcategory:\"文件系统\"、auditpol /set /subcategory:\"文件系统\" /success:en；3.2.2 中标麒麟、银河麒麟：sudo bash -c 'echo \"blacklist usb-storage\" >>、sudo firew" ;;
        3.3) printf %s "3.3.1 Win7\\WinXP：where cipher、cipher /w:D:、cipher /w:E:；3.3.2 中标麒麟、银河麒麟：sudo apt-get install coreutils、sudo shred -n 3 -z -v /dev/sdb；3.3.3 通用核查方法（专业擦除工具层）：图形界面/文档核查" ;;
        3.4) printf %s "3.4.1 通用核查方法（物理销毁）：图形界面/文档核查" ;;
        3.5) printf %s "3.5.1 Win7\\WinXP：icacls C:\\Windows\\System32\\winevt\\Logs、sc qc eventlog、manage-bde -status；3.5.2 中标麒麟、银河麒麟：sudo visudo、audit_admin ALL=(ALL) NOPASSWD: /bin/cat, /bi、sudo apt-get install logrotate" ;;
        3.6) printf %s "3.6.1 边界设备核查方法：图形界面/文档核查；3.6.2 合规判定：图形界面/文档核查；3.6.3 通用核查方法（DLP 系统层）：图形界面/文档核查" ;;
        3.7) printf %s "3.7.1 Win7\\WinXP：以管理员身份运行 Windows Terminal（或命令提示符）、net user、auditpol /get /category:*；3.7.2 中标麒麟、银河麒麟：sudo storage-cli login、sudo storage-cli user list、sudo storage-cli audit log list --type=manage" ;;
        3.8) printf %s "3.8.1 Win7、WinXP：以管理员身份运行 Windows Terminal（或命令提示符）、diskpart、list disk；3.8.2 中标麒麟、银河麒麟：sudo apt-get update、sudo apt-get install suricata -y、alert http any any -> any any (msg:\"DLP: 检测到'" ;;
        3.9) printf %s "3.9.1 Win7、WinXP：以管理员身份运行 Windows Terminal（或命令提示符）、Windows 日志–安全；3.9.2 中标麒麟、银河麒麟：sudo apt-get install libpam-google-authentica、sudo -u 用户名 google-authenticator、sudo vi /etc/pam.d/sshd" ;;
        3.10) printf %s "3.10.1 Win7、WinXP：Windows 日志–安全；3.10.2 中标麒麟、银河麒麟：图形界面/文档核查；3.10.3 通用核查方法（应用层与网络层）：图形界面/文档核查" ;;
        3.11) printf %s "3.11.1 Win7、WinXP：Windows 日志–安全；3.11.2 中标麒麟、银河麒麟：grep -r \"CryptoPolicy\" /etc/ssh/sshd_config /、ip route show、sudo find /var/log -name \"*network*\" -o -name；3.11.3 补充核查方法（数据库层与应用层权限分级）：图形界面/文档核查" ;;
        3.12) printf %s "3.12.1 Win7、WinXP：图形界面/文档核查；3.12.2 中标麒麟、银河麒麟：sudo smbstatus、sudo netstat -tulpn | grep -E \":139|:445|:204；3.12.3 通用核查方法（应用层）：图形界面/文档核查" ;;
        3.13) printf %s "3.13.1 Win7、WinXP：Windows 日志–安全；3.13.2 中标麒麟、银河麒麟：图形界面/文档核查；3.13.3 通用核查方法（应用层与数据库层）：图形界面/文档核查" ;;
        3.14) printf %s "3.14.1 Win7、WinXP：Windows 日志–安全；3.14.2 中标麒麟、银河麒麟：sudo find / -type d -name \"*l[1-4]*\" -exec ls、[ -f \"/etc/krb5.conf\" ] && echo 是 || echo 否；3.14.3 补充核查方法（大数据平台层统一管控）：图形界面/文档核查" ;;
        4.1) printf %s "4.1.1 Windows操作系统 (Win7, WinXP)：dir C:\\ /ad /s | findstr /i \"backup bak、dir D:\\backup /o:-d、schtasks；4.1.2 中标麒麟、银河麒麟：图形界面/文档核查；4.1.3 补充核查方法（应用层与数据库层）：备份/恢复" ;;
        4.2) printf %s "4.2.1 Windows操作系统 (Win7, WinXP)：wmic qfe list brief /format:table、type C:\\Windows\\WindowsUpdate.log、Get-ChildItem \"C:\\App\\Logs\" -Filter \"*update*" ;;
        4.3) printf %s "4.3.1 Windows操作系统 (Win7, WinXP)：sigcheck.exe -q -m C:\\路径\\到\\程序.exe、sigcheck.exe -q -m C:\\Windows\\System32\\notepa、Get-AuthenticodeSignature \"C:\\路径\\到\\程序.exe\"" ;;
        4.4) printf %s "4.4.1 Windows操作系统 (Win7, WinXP)：netstat -ano | findstr LISTENING、tasklist /FI \"PID eq <PID号>\"、sc query state= all；4.4.2 中标麒麟、银河麒麟：sudo netstat -tulpn | grep LISTEN | awk '{pri、sudo iptable" ;;
        4.5) printf %s "4.5.1 Windows操作系统 (Win7, WinXP)：netsh advfirewall show allprofiles、netsh advfirewall firewall show rule name=all" ;;
        4.6) printf %s "4.6.1 通用核查方法：图形界面/文档核查；4.6.2 主要判定标准：图形界面/文档核查；4.6.3 通用核查方法（数据库连接最小权限补充）：图形界面/文档核查" ;;
        4.7) printf %s "4.7.1 Win7、WinXP：Select-String -Path C:\\Windows\\System32\\inets" ;;
        4.8) printf %s "4.8.1 Win7、WinXP：图形界面/文档核查；4.8.2 中标麒麟、银河麒麟：sudo grep -E \"/var/www|html\" /etc/aide/aide.c、sudo ls -la /var/www/html/index.*；4.8.3 补充核查方法（专业防篡改系统与文件完整性）：图形界面/文档核查" ;;
        4.9) printf %s "4.9.1 Win7、WinXP：图形界面/文档核查；4.9.2 中标麒麟、银河麒麟：curl -I http://目标URL | grep -iE \"(X-Content-T、curl -s \"$URL' OR '1'='1\" | grep -i \"error\\|s；4.9.3 通用核查方法（应用层SQL注入与XSS防护）：图形界面/文档核查" ;;
        4.10) printf %s "4.10.1 Win7、WinXP：图形界面/文档核查；4.10.2 中标麒麟、银河麒麟：nikto -h http://192.168.1.100 -Tuning 1,3,5 -、curl -I http://192.168.1.100:8080/webapp/；4.10.3 通用核查方法（应用层执行代码验证）：图形界面/文档核查" ;;
        4.11) printf %s "4.11.1 Win7、WinXP：图形界面/文档核查；4.11.2 中标麒麟、银河麒麟：图形界面/文档核查；4.11.3 补充核查方法（应用层RBAC与数据库授权）：图形界面/文档核查" ;;
        4.12) printf %s "4.12.1 Win7、WinXP：1..200 | ForEach-Object {

    Start-Job { In、Get-Job；4.12.2 中标麒麟、银河麒麟：cat /proc/sys/fs/file-max、ss -s | grep 'TCP:' | awk '{print $2}'；4.12.3 通用核查方法（应用层与Web服务器层）：图形界面/文档核查" ;;
        4.13) printf %s "4.13.2 中标麒麟、银河麒麟：sudo grep -E \"(Accepted|Failed)\" /var/log/aut；4.13.3 通用核查方法（业务管理终端专设专用）：图形界面/文档核查" ;;
        4.14) printf %s "4.14.1 Win7、WinXP：图形界面/文档核查；4.14.2 中标麒麟、银河麒麟：图形界面/文档核查；4.14.3 通用核查方法（应用层与数据库层细粒度授权）：图形界面/文档核查" ;;
        4.15) printf %s "4.15.1 Win7、WinXP：图形界面/文档核查；4.15.2 中标麒麟、银河麒麟：sudo sestatus -v 2>/dev/null | grep -i \"mls\\|、sudo ls -Z /etc/passwd 2>/dev/null；4.15.3 通用核查方法（应用层签名验证与密级标识）：图形界面/文档核查" ;;
        4.16) printf %s "4.16.1 Win7、WinXP：图形界面/文档核查；4.16.2 中标麒麟、银河麒麟：图形界面/文档核查；4.16.3 通用核查方法（应用层与数据库层远程管理加密）：图形界面/文档核查" ;;
        4.17) printf %s "4.17.1 Win7、WinXP：图形界面/文档核查；4.17.2 中标麒麟、银河麒麟：图形界面/文档核查；4.17.3 通用核查方法（应用层与数据库层三权分立）：图形界面/文档核查" ;;
        4.18) printf %s "4.18.1 Win7、WinXP：图形界面/文档核查；4.18.2 中标麒麟、银河麒麟：图形界面/文档核查；4.18.3 通用核查方法（应用层与网络层地址限制）：图形界面/文档核查" ;;
        4.19) printf %s "4.19.1 Win7、WinXP：图形界面/文档核查；4.19.2 中标麒麟、银河麒麟：图形界面/文档核查；4.19.3 补充核查方法（应用层登录失败处理）：图形界面/文档核查" ;;
        4.20) printf %s "4.20.1 Win7、WinXP：图形界面/文档核查；4.20.2 中标麒麟、银河麒麟：sudo systemctl list-unit-files | grep -i \"kyl、sudo find /etc -name \"*.conf\" -type f | xargs、sudo systemctl list-units | grep -i \"gmssl\\|s" ;;
        4.21) printf %s "4.21.1 Win7、WinXP：图形界面/文档核查；4.21.2 中标麒麟、银河麒麟：sudo grep -r \"pam_fprintd\\|pam_biometric\" /et、openssl ecparam -list_curves 2>/dev/null | gr；4.21.3 通用核查方法（应用层数字证书认证）：图形界面/文档核查" ;;
        4.22) printf %s "4.22.2 中标麒麟、银河麒麟：sudo netstat +-tunlp；4.22.3 通用核查方法（更改Web应用系统默认服务发布端口）：图形界面/文档核查" ;;
        4.23) printf %s "4.23.1 Win7、WinXP：图形界面/文档核查；4.23.2 中标麒麟、银河麒麟：图形界面/文档核查；4.23.3 通用核查方法（Web服务器层与数据库层管理端口分离）：图形界面/文档核查" ;;
        4.24) printf %s "4.24.1 Win7、WinXP：图形界面/文档核查；4.24.2 中标麒麟、银河麒麟：图形界面/文档核查；4.24.3 通用核查方法（数据库连接地址核查）：图形界面/文档核查" ;;
        4.25) printf %s "4.25.1 Win7、WinXP：图形界面/文档核查；4.25.2 中标麒麟、银河麒麟：sudo systemctl status auditd 2>/dev/null || s、ls -lh /var/log/auth.log 2>/dev/null、grep -r \"rotate\\|maxage\\|180\" /etc/logrotate." ;;
        4.26) printf %s "4.26.1 Win7、WinXP：图形界面/文档核查；4.26.2 中标麒麟、银河麒麟：find /var/log /opt /root /home -name \"*审计*\" -、grep -i \"CVE\\|漏洞\\|修复\\|patch\" /var/log/yum.log；4.26.3 通用核查方法（应用层SAST/DAST与代码审计流程）：图形界面/文档核查" ;;
        4.27) printf %s "4.27.1 Win7、WinXP：图形界面/文档核查；4.27.2 中标麒麟、银河麒麟：图形界面/文档核查；4.27.3 通用核查方法（应用层前端与后端双重校验）：图形界面/文档核查" ;;
        4.28) printf %s "4.28.1 Win7、WinXP：图形界面/文档核查；4.28.2 中标麒麟、银河麒麟：curl -I http://目标URL | grep -iE \"(X-Content-T、curl -s \"$URL' OR '1'='1\" | grep -i \"error\\|s；4.28.3 通用核查方法（应用层四类攻击防御）：图形界面/文档核查" ;;
        4.29) printf %s "4.29.1 Win7、WinXP：图形界面/文档核查；4.29.2 中标麒麟、银河麒麟：图形界面/文档核查；4.29.3 补充核查方法（应用层统一权限管理平台）：图形界面/文档核查" ;;
        4.30) printf %s "4.30.1 Win7、WinXP：图形界面/文档核查；4.30.2 中标麒麟、银河麒麟：图形界面/文档核查；4.30.3 通用核查方法（应用层管理面限制）：图形界面/文档核查" ;;
        4.31) printf %s "4.31.1 Win7、WinXP：图形界面/文档核查；4.31.2 中标麒麟、银河麒麟：grep -r \"maxlogins\\|maxsyslogins\" /etc/securi；4.31.3 通用核查方法（应用层与Web服务器层并发会话）：图形界面/文档核查" ;;
        4.32) printf %s "4.32.1 Win7、WinXP：图形界面/文档核查；4.32.2 中标麒麟、银河麒麟：find /opt /etc /root -name \"*备份*\" -o -name \"*；4.32.3 补充核查方法（应用层与数据库层备份恢复）：图形界面/文档核查" ;;
        4.33) printf %s "4.33.2 中标麒麟、银河麒麟：cat /etc/os-release | grep -E \"NAME|VERSION|I、uname -a、dpkg -l 2>/dev/null | grep -v \"ubuntu\\|debian；4.33.3 通用核查方法（国产自主可控软硬件）：图形界面/文档核查" ;;
        5.1) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.2) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.3) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.4) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.5) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.6) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.7) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.8) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.9) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.10) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.11) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.12) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.13) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.14) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.15) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.16) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.17) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.18) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.19) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.20) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.21) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.22) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        5.23) printf %s "采集-解析：按采集清单登录设备执行只读命令、保存回显，脚本解析回显判定（命令清单见 check_network 采集模板）" ;;
        6.1) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.2) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.3) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.4) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.5) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.6) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.7) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.8) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        6.9) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        7.1) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        7.2) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        7.3) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        7.4) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        7.5) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        8.1) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        8.2) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        8.3) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        8.4) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        9.1) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        9.2) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        9.3) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        9.4) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        10.1) printf %s "人工核查（现场查看 / 文档调阅，按指导书该条方法留存证据）" ;;
        *) printf %%s '' ;;
    esac
}

guide_title_of() {
    case "$1" in
        1.1) printf %s "操作系统、数据库管理系统、中间件等平台软件应及时安装补丁程序" ;;
        1.2) printf %s "操作系统应安装防病毒软件并及时升级" ;;
        1.3) printf %s "操作系统应按需求裁剪服务和端口" ;;
        1.4) printf %s "操作系统应具备防火墙功能" ;;
        1.5) printf %s "操作系统应停用冗余网络设置" ;;
        1.6) printf %s "操作系统远程管理应开放唯一管理服务,指定管理终端并采取传输加密保护措施" ;;
        1.7) printf %s "数据库管理系统应删除冗余帐户,应设置不少于8个字符且字符采用字母大小写，数字及特殊字符混合编制的账户口令" ;;
        1.8) printf %s "数据库管理系统应删除冗余存储过程" ;;
        1.9) printf %s "数据库管理系统应具有基于表级增删改查等细粒度访问和管理授权功能" ;;
        1.10) printf %s "数据库管理系统应具有自主访问控制功能" ;;
        1.11) printf %s "数据库管理系统应具有备份和恢复功能" ;;
        1.12) printf %s "数据库管理系统应具有表级审计、告警和阻断功能" ;;
        1.13) printf %s "数据库管理系统的数据应和其他应用的数据分类独立存储" ;;
        1.14) printf %s "中间件应采取限制运行权限和使用安全管理通道等安全加固措施" ;;
        1.15) printf %s "应具备数据库管理系统超级管理员远程登录限制远程登陆限制能力" ;;
        1.16) printf %s "应具备数据库管理系统输入（参数）检查能力" ;;
        1.17) printf %s "应限制操作系统开放的远程管理服务或端口" ;;
        1.18) printf %s "应限制用户对服务器资源的最大或最小使用限度" ;;
        1.19) printf %s "应更换数据库管理系统的默认服务端口、管理员用户名和口令" ;;
        1.20) printf %s "数据库管理系统应配置安全策略" ;;
        1.21) printf %s "数据库管理系统应具有行级或列级审计功能" ;;
        1.22) printf %s "数据库管理系统应采取单独、安全监控、审计措施" ;;
        1.23) printf %s "数据库管理系统仅为应用服务器提供访问服务" ;;
        1.24) printf %s "应具备日志审计能力，审计日志至少保留180天" ;;
        1.25) printf %s "检查是否具备边界保护能力,是否可以抗攻击，防纂改" ;;
        1.26) printf %s "检查是否有防病毒日志、补丁日志、记录相关信息的完整，有效" ;;
        1.27) printf %s "操作系统、数据库管理系统、中间件、办公软件等基础软件应采用具有完备的售后技术支持与服务的正版或定制软件" ;;
        2.1) printf %s "服务器和用户计算机应设置登录口令" ;;
        2.2) printf %s "用户计算机应根据需要安装补丁程序" ;;
        2.3) printf %s "用户应设置用户应用口令，通过认证后使用信息服务" ;;
        2.4) printf %s "用户计算机应具有互不相同的用户名和口令" ;;
        2.5) printf %s "用户计算机应关闭冗余系统服务和端口" ;;
        2.6) printf %s "用户计算机应具备阻断和告警非法连接互联网的能力" ;;
        2.7) printf %s "用户计算机应具有文件保护功能" ;;
        2.8) printf %s "应采取终端管控措施、具有统一配置、安全加固、网络访问控制、外设接口管控、软件进程管控、无线模块禁用、防止IP地址和MAC地址非授权改动等功能" ;;
        2.9) printf %s "用户计算机应禁止安装与工作无关的软件" ;;
        2.10) printf %s "用户计算机USB接口应禁止私自连接对拷线和手机、媒体播放设备等个人移动电子设备" ;;
        2.11) printf %s "用户计算机登录应使用基于专用物理部件或生物特征的多因素身份认证方式，应设置超时锁屏，屏幕保护等待时间不超过5min,服务器应设置登录口令，口令长度不得少于10个字符，字符应采用字母大小写、数字及特殊字符混合编制，更换周期不超过30d" ;;
        2.12) printf %s "用户计算机应物理拆除Wi-Fi、红外、蓝牙等无线模块，确需使用Wi-Fi的应严格采用JY密码等措施保护" ;;
        2.13) printf %s "用户计算机应采取非法外联阻断、文件输出管控等措施" ;;
        2.14) printf %s "应具备用户行为审计能力；审计日志留存期应满足制度要求" ;;
        2.15) printf %s "检查安全策略配置情况和设置功能" ;;
        2.16) printf %s "检查被试装备中操作系统、数据库以及应用软件等是否完成补丁修复和升级到最新版本" ;;
        3.1) printf %s "集中存储的涉密数据应采取加密保护措施" ;;
        3.2) printf %s "用户计算机之间的数据交互、文件传输均应统一管控和审计" ;;
        3.3) printf %s "涉密存储载体在降密级使用前或重大JS演训活动结束后，应采取数据写覆盖方法及时清除数据" ;;
        3.4) printf %s "对确定销毁的涉密载体，应采取消磁、粉碎、溶解、化浆和熔化等方法进行销毁" ;;
        3.5) printf %s "网络、系统、应用和用户行为等日志应采取读写控制、加密、变换、完整性校验等保护措施，TM级数据存储1应采取加密保护措施" ;;
        3.6) printf %s "网络边界应通过边界设备具备信息过滤、敏感内容识别等数据防泄漏能力" ;;
        3.7) printf %s "数据存储系统管理登录至少采取验证码等增强措施，修改默认用户名和口令等默认设置，管理与访问应具有行为审计功能，审计日志应至少保留180天" ;;
        3.8) printf %s "数据存储系统应根据重要程度划分不同存储区块，并设置用户访问权限" ;;
        3.9) printf %s "检查数据传输过程是否按照要求进行加密，传输路径是否合理，是否统一管控、留有日志记录，是否具有防泄漏措施，是否存在安全风险" ;;
        3.10) printf %s "检查数据共享是否合理，是否存在安全隐患" ;;
        3.11) printf %s "检查对数据的访问是否按照权限分级访问，是否对访问行为进行审计" ;;
        3.12) printf %s "检查数据采集是否超出业务需求范围" ;;
        3.13) printf %s "检查数据各环节处理是否满足密级相应的保密要求" ;;
        3.14) printf %s "检查对数据的访问是否按照权限分级访问，是否对访问行为进行审计，对大数据的访问是否提供同一管控和访问控制" ;;
        4.1) printf %s "应用系统软件应采取备份措施" ;;
        4.2) printf %s "应用系统软件应及时安装补丁程序，且更新所用升级包应经过安全性测试" ;;
        4.3) printf %s "基于可信根对应用系统软件进行可信验证。可信性受到破坏后报警" ;;
        4.4) printf %s "提供公共信息服务的服务器应与涉密信息服务器分设，专用服务器应只提供专用服务" ;;
        4.5) printf %s "提供公共信息服务的服务器应具备防DDoS攻击能力" ;;
        4.6) printf %s "Web应用系统应采取Web安全防护措施" ;;
        4.7) printf %s "网站服务宜以静态页面形式发布" ;;
        4.8) printf %s "网站应采取网页防篡改措施，防止对信息内容的非法修改" ;;
        4.9) printf %s "Web应用系统应具备防范SQL注入、跨站脚本等攻击能力" ;;
        4.10) printf %s "Web应用系统应具备执行代码有效验证能力" ;;
        4.11) printf %s "应具备用户授权访问控制能力" ;;
        4.12) printf %s "应具备访问应用最大并发会话连接数限制能力" ;;
        4.13) printf %s "业务管理终端专设专用" ;;
        4.14) printf %s "应具有基于用户角色的授权访问控制能力，访问控制主体的细粒度达到用户级或进程级，客体的细粒度应达到文件级、表和记录级、字段级" ;;
        4.15) printf %s "文电等文档专用业务处理系统应具有签名验证、密级标识等功能" ;;
        4.16) printf %s "远程管理应采取加密保护措施" ;;
        4.17) printf %s "应支持管理员、安全员、审计员三权分立的职责划分，禁止设立超级管理员，并限制管理员、安全员、审计员的数量" ;;
        4.18) printf %s "业务管理终端登录应采取网络地址限制措施" ;;
        4.19) printf %s "应具有结束会话、限定登录错误次数和自动退出等登录失败处理功能" ;;
        4.20) printf %s "宜使用自主设计开发的网络服务、协议、接口等，增强应用安全" ;;
        4.21) printf %s "应具有基于专用物理部件或生物特征多因素的、与JD密码算法相结合的数字证书用户身份认证功能" ;;
        4.22) printf %s "应更改Web应用系统默认服务发布端口" ;;
        4.23) printf %s "应分开设置管理端口与应用端口" ;;
        4.24) printf %s "应用服务和数据存储应部署在不同的服务器上" ;;
        4.25) printf %s "应具有对所有访问行为和管理行为的日志审计功能，支持多组合查询检索，可读性强，具有解释和展示功能，审计日志应至少保留180天" ;;
        4.26) printf %s "应经过代码级安全漏洞挖掘" ;;
        4.27) printf %s "检查是否具备对人机接口输入、网络通信输入、文件输入的数据进行格式和长度检查的功能" ;;
        4.28) printf %s "检查是否能够有效检测并防御SQL注入、网页篡改、跨站脚本、拒绝服务等应用层攻击" ;;
        4.29) printf %s "检查是否具有用户访问权限统一管理功能" ;;
        4.30) printf %s "检查是否具备统一管理措施，是否对远程管理进行限制" ;;
        4.31) printf %s "检查是否能够设置最大并发会话连接数、会话建立速率、单用户并发会话数" ;;
        4.32) printf %s "检查所有应用是否具备备份与恢复功能" ;;
        4.33) printf %s "检查应用软件是否基于国产自主可控软硬件自主开发" ;;
        5.1) printf %s "JD网络跨网跨域数据交换时应按规定流程进行" ;;
        5.2) printf %s "利用无线网络技术构建高防护等级网络时应按规定流程执行" ;;
        5.3) printf %s "应按最小化原则设计网络架构" ;;
        5.4) printf %s "网络边界物理互联节点、路由协议与路由地址网段应满足最小化原则" ;;
        5.5) printf %s "局域网内部应根据业务性质划分安全区域" ;;
        5.6) printf %s "网络设备应按最小化原则进行远程管理、账户设置、访问控制等安全配置" ;;
        5.7) printf %s "网络设备应开放唯一网络管理服务并限制管理终端访问" ;;
        5.8) printf %s "远程管理网络设备与安全防护设备应采用加密保护的管理服务" ;;
        5.9) printf %s "同链路相同安全功能的防护设备应使用不同架构或不同品牌" ;;
        5.10) printf %s "应对组播源、组播地址、组播成员采取控制措施" ;;
        5.11) printf %s "重要网络设备和网络安全防护设备应有备份" ;;
        5.12) printf %s "远程传输应采取两层加密保护措施" ;;
        5.13) printf %s "应设立网络安全管理中心并指定管理终端" ;;
        5.14) printf %s "与互联网等外部网络应按等级采取物理或逻辑隔离" ;;
        5.15) printf %s "不同用途网络之间应通过防护设备加强逻辑隔离并具备攻击告警审计阻断能力" ;;
        5.16) printf %s "局域网各安全区域之间及主机之间应采取全网细粒度访问控制" ;;
        5.17) printf %s "远程租用线路传输应采取三层加密保护措施" ;;
        5.18) printf %s "应采取不低于802.1x认证强度的接入认证措施" ;;
        5.19) printf %s "用户计算机之间应逻辑隔离" ;;
        5.20) printf %s "应具备全网行为审计记录能力且审计日志至少保留180天" ;;
        5.21) printf %s "应对全网攻击与违规行为采取实时监视报警审计控制阻断定位措施" ;;
        5.22) printf %s "数据存储系统的管理网络与应用网络应逻辑隔离" ;;
        5.23) printf %s "检查网络设备配备的合理性与必要性" ;;
        6.1) printf %s "使用的网络设备、服务器、终端等，应选用进入《全J计算机及网络设备集中采购目录》的产品" ;;
        6.2) printf %s "使用的安全网关、防火墙等信息安全产品，应通过JD信息安全测评认证机构的认证" ;;
        6.3) printf %s "机房应有序、规范、合理走线布线，明确互联网区域和内部区域，粘贴标识，区分不同线路" ;;
        6.4) printf %s "机房应符合GB 2887-2011中4.6.1所规定的温湿度等要求" ;;
        6.5) printf %s "机房应安装安防监控设备，对机房的人员进出、设备操作等进行监管" ;;
        6.6) printf %s "机房应将涉密区域和互联网区域设置于不同场所" ;;
        6.7) printf %s "在涉密网络中使用过的打印机、复印机、刻录机、扫描仪、存储载体等设备严禁在互联网中使用" ;;
        6.8) printf %s "应选用《JY关键软硬件自主可控产品名录》中的芯片类、计算机及外设类、网络设备类、安全防护设备类、存储设备类等产品" ;;
        6.9) printf %s "与安全防护等级低的网络使用不同色系线缆进行严格隔离区分" ;;
        7.1) printf %s "应设置安全管理机构，保证系统安全措施的落实" ;;
        7.2) printf %s "应配备专职系统安全保密管理人员，负责系统安全措施的落实" ;;
        7.3) printf %s "应具有安全管理领导机构，督导系统安全措施的落实" ;;
        7.4) printf %s "应具有系统安全保密技术管理人员，具体负责安全保密技术措施的落实，保证系统安全运行" ;;
        7.5) printf %s "应具有完善的应急响应体系，应对突发事件" ;;
        8.1) printf %s "应具有日常安全管理制度、入网审批制度和系统安全保密检查制度等" ;;
        8.2) printf %s "应具有安全管理操作规程、安全监控操作规程、安全审计操作规程、应急响应操作规程等" ;;
        8.3) printf %s "应具有日常系统备份制度和存储载体使用管理制度" ;;
        8.4) printf %s "应具有脆弱性分析等规程" ;;
        9.1) printf %s "应具有应急响应预案，当发生危及系统安全的事件时应根据操作规程及时采取措施，符合应急预案启动条件时按预案开展应急措施" ;;
        9.2) printf %s "计算机信息系统安全保密管理人员应能对网络中的网络安全防护设备进行统一配置、管理，并能实现安全管理中心与网络安全防护设备之间的响应" ;;
        9.3) printf %s "应具有与实际情况相符且完整的安全保密策略文档和安全保密技术及产品配置的详细记录" ;;
        9.4) printf %s "安全保密技术管理人员应对每日网络和系统运行情况实施安全审计，每月组织本级网络的安全性检测，编写安全审计与评估报告，并形成记录，系统配置发生变化的情况下应及时组织安全检测评估" ;;
        10.1) printf %s "审计协议的机密性、完整性、认可性、不可否认性" ;;
        *) printf %%s '' ;;
    esac
}

add_result() {
    # 评审整改（2026-09-30 第4/8条）：R_METHOD=验证过程/方法（指导书逐条提取）；
    # 核查项名对齐指导书（不一致时以指导书 H2 标题为准，原叫法并入详情前缀）
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"; R_CAT[$R_COUNT]="$2"; R_STATUS[$R_COUNT]="$4"
    R_CHAPTER[$R_COUNT]="$6"; R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="$(guide_ref "$1")"
    local gtitle; gtitle="$(guide_title_of "$1")"
    if [ -n "$gtitle" ] && [ "$gtitle" != "$3" ]; then
        R_TITLE[$R_COUNT]="$gtitle"
        R_DETAIL[$R_COUNT]="〔脚本项名：$3〕$5"
    else
        R_TITLE[$R_COUNT]="$3"
        R_DETAIL[$R_COUNT]="$5"
    fi
    R_METHOD[$R_COUNT]="$(method_of "$1")"
}

db_unreachable() {
    if [ "$DM_PRESENT" = "0" ]; then
        add_result "$1" "系统安全-数据库" "$2" "na" "未检测到 disql 客户端，无法对达梦数据库进行自动核查，标记不适用。" "第1章" "$3"
    else
        add_result "$1" "系统安全-数据库" "$2" "manual" "无法连接达梦数据库（${DM_USER}@${DM_HOST}:${DM_PORT}），需人工登录核查。请检查 DM_USER/DM_PASS 环境变量。" "第1章" "$3"
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
<td>$(html_esc "${R_METHOD[$i]}")</td>
<td>$(html_esc "${R_DETAIL[$i]}")</td>
<td>$(html_esc "$([ "${R_STATUS[$i]}" = fail ] && printf '%s' "${R_REC[$i]}")")</td>
<td>$(html_esc "$([ "${R_STATUS[$i]}" != fail ] && printf '%s' "${R_REC[$i]}")")</td>
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
        printf '{"ch":"%s","id":"%s","cat":"%s","title":"%s","status":"%s","method":"%s","detail":"%s","rec":"%s","guide":"%s"}' \
            "$(json_esc "${R_CHAPTER[$i]}")" "$(json_esc "${R_ID[$i]}")" "$(json_esc "${R_CAT[$i]}")" \
            "$(json_esc "${R_TITLE[$i]}")" "$(json_esc "${R_STATUS[$i]}")" "$(json_esc "${R_METHOD[$i]}")" "$(json_esc "${R_DETAIL[$i]}")" \
            "$(json_esc "${R_REC[$i]}")" "$(json_esc "${R_GUIDE[$i]}")"
    done
    printf ']'
}

REPORT_TAG="达梦"
REPORT_TITLE="达梦数据库配置核查报告"
report_meta_html() {
    echo "<div>连接：$(html_esc "$DM_USER@$DM_HOST:$DM_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")</div>"
    echo "<div>主机名：$(hostname 2>/dev/null)　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>"
    echo "<div>参考标准：配置核查作业指导书v2.0.0 / 配置核查表_v2.0.0.xlsx</div>"
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
  <div class="sub">参考标准：配置核查作业指导书v2.0.0</div>
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
<thead><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr></thead>
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
      "<td class='method'>"+esc(x.method)+"</td>"+
      "<td class='detail'>"+esc(x.detail)+"</td>"+
      "<td class='rec'>"+(x.status=='fail'?esc(x.rec):"")+"</td>"+
      "<td class='req'>"+(x.status=='fail'?"":esc(x.rec))+"</td>"+
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
    local outfile="$OUT_DIR/配置核查报告_达梦_${STAMP}.xls"
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
<x:Name>达梦配置核查报告</x:Name>
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
<p>达梦数据库配置核查报告　连接：$(html_esc "$DM_USER@$DM_HOST:$DM_PORT")　版本：$(html_esc "${DB_VERSION:-未连接}")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>
XLSHEAD
        if [ "$fail" -gt 0 ]; then echo "<p><b>一、未通过（$fail 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>"; build_table_rows_by_status fail; echo "</table>"; fi
        if [ "$manual" -gt 0 ]; then echo "<p><b>二、需人工核查（$manual 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>"; build_table_rows_by_status manual; echo "</table>"; fi
        if [ "$na" -gt 0 ]; then echo "<p><b>三、不适用（$na 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>"; build_table_rows_by_status na; echo "</table>"; fi
        if [ "$pass" -gt 0 ]; then echo "<p><b>四、通过（$pass 项）</b></p><table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>"; build_table_rows_by_status pass; echo "</table>"; fi
        cat <<XLSFOOT
</body>
</html>
XLSFOOT
    } > "$outfile"
    echo "Excel(.xls)报告已生成：$outfile"
}

# ============================================================
# 核查项（共19项，对应配置核查表 达梦 列）
# ============================================================

check_1_1_patch() {
    add_result "1.1" "系统安全-数据库" "数据库补丁程序" "manual" "达梦当前版本：${DB_VERSION:-未知}。请人工核对是否为最新补丁版本（对比达梦官方版本发布）。" "第1章" "及时升级到最新补丁版本。"
}

check_1_7_dbaccounts() {
    local pwd_policy
    pwd_policy="$(dm_ini PWD_POLICY)"
    if [ -z "$pwd_policy" ]; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "manual" "未读取到 PWD_POLICY 参数，请人工核查口令策略与冗余账户（SYSDBA/SYSAUDITOR/SYSSSO 等默认账户是否加固）。" "第1章" "设置 PWD_POLICY 口令复杂度，删除/锁定冗余账户。"
    elif [ "$pwd_policy" -lt 15 ] 2>/dev/null; then
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "fail" "口令策略 PWD_POLICY=${pwd_policy}（<15，未启用完整口令复杂度：长度/大小写/数字/特殊字符）。" "第1章" "设置 PWD_POLICY>=15，启用完整口令复杂度策略。"
    else
        add_result "1.7" "系统安全-数据库" "数据库账户管理" "pass" "口令策略 PWD_POLICY=${pwd_policy}（已启用完整口令复杂度）。" "第1章" "持续核查并删除/锁定冗余账户，保持强口令策略。"
    fi
}

check_1_8_dbprocedures() {
    local procs
    procs="$(dm_count "SELECT COUNT(*) FROM DBA_PROCEDURES;")"
    if [ -n "$procs" ]; then
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "manual" "存储过程/函数总数：${procs}。请人工确认是否存在冗余或高危存储过程。" "第1章" "删除冗余存储过程，遵循最小功能原则。"
    else
        add_result "1.8" "系统安全-数据库" "数据库存储过程管理" "manual" "未读取到存储过程数量，请人工核查冗余存储过程。" "第1章" "删除冗余存储过程。"
    fi
}

check_1_9_dbgrants() {
    local tab_priv
    tab_priv="$(dm_count "SELECT COUNT(*) FROM DBA_TAB_PRIVS;")"
    if [ -n "$tab_priv" ] && [ "$tab_priv" -gt 0 ] 2>/dev/null; then
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "pass" "存在对象级（表/列）授权 ${tab_priv} 条，已按细粒度授权。" "第1章" "持续按最小权限原则，使用表级/列级授权。"
    else
        add_result "1.9" "系统安全-数据库" "数据库权限最小化" "manual" "未发现对象级细粒度授权记录，请人工确认是否已按最小权限原则分配账户权限。" "第1章" "按最小权限原则分配账户权限。"
    fi
}

check_1_10_dbaccess() {
    local listen
    listen="$(ss -lntp 2>/dev/null | awk -v p=":$DM_PORT" '$4 ~ (":"p"$"){print $4}')"
    if echo "$listen" | grep -qE '^(0\.0\.0\.0|\*|\[::\])' 2>/dev/null; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "fail" "数据库端口监听在 0.0.0.0/::（对外暴露）：$listen" "第1章" "将监听地址改为内网/127.0.0.1，配合防火墙限制访问来源。"
    elif [ -n "$listen" ]; then
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "pass" "数据库未对外暴露监听：$listen。" "第1章" "持续限制数据库监听地址与访问来源。"
    else
        add_result "1.10" "系统安全-数据库" "数据库访问控制" "manual" "未在默认端口 ${DM_PORT} 检测到监听，可能使用自定义端口，需人工确认监听地址。" "第1章" "确认数据库实际监听端口与地址。"
    fi
}

check_1_11_dbbackup() {
    local cron
    cron="$(grep -ril 'dmrman\|dmbackup\|BACKUP[[:space:]]*DATABASE' /etc/cron.d /etc/cron.daily /var/spool/cron 2>/dev/null)"
    if [ -n "$cron" ]; then
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "pass" "检测到达梦备份相关计划任务：$cron" "第1章" "定期验证备份可恢复性，并异地存储。"
    else
        add_result "1.11" "系统安全-数据库" "数据库备份策略" "manual" "未在常见 crontab 位置发现达梦备份任务（dmrman/BACKUP），请人工核实备份策略。" "第1章" "建立定期备份（dmrman/BACKUP DATABASE）。"
    fi
}

check_1_12_dbaudit() {
    local enable_audit
    enable_audit="$(dm_ini ENABLE_AUDIT)"
    if [ "$enable_audit" = "1" ]; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "pass" "达梦审计已开启（ENABLE_AUDIT=${enable_audit:-1}）。" "第1章" "确保审计覆盖登录、权限变更、敏感数据操作。"
    elif [ -n "$enable_audit" ]; then
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "fail" "达梦审计未开启（ENABLE_AUDIT=${enable_audit}）。" "第1章" "开启数据库审计功能。"
    else
        add_result "1.12" "系统安全-数据库" "数据库审计插件" "manual" "未读取到 ENABLE_AUDIT 参数，请人工确认审计是否开启。" "第1章" "开启数据库审计功能。"
    fi
}

check_1_13_dbisolation() {
    add_result "1.13" "系统安全-数据库" "数据库与业务隔离" "manual" "数据分类独立存储需结合业务架构人工核查（分库/分表/分实例存储）。" "第1章" "敏感数据与普通数据分类独立存储。"
}

check_1_15_dbremote() {
    local listen
    listen="$(ss -lntp 2>/dev/null | awk -v p=":$DM_PORT" '$4 ~ (":"p"$"){print $4}')"
    if echo "$listen" | grep -qE '^(0\.0\.0\.0|\*|\[::\])' 2>/dev/null; then
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "fail" "数据库端口对外暴露：$listen，SYSDBA 等超管账户存在被远程访问风险。" "第1章" "限制监听地址与防火墙规则，禁止超管远程登录。"
    else
        add_result "1.15" "系统安全-数据库" "数据库远程访问控制" "manual" "数据库未对外暴露监听，达梦默认限制 SYSDBA 远程登录，请人工确认超管远程登录限制。" "第1章" "限制超管远程登录来源。"
    fi
}

check_1_16_inputcheck() {
    add_result "1.16" "系统安全-数据库" "输入验证/SQL注入防护" "manual" "数据库输入（参数）检查属应用层能力，需人工/工具核查应用是否使用参数化查询防 SQL 注入。" "第1章" "应用层使用参数化查询。"
}

check_1_19_dbdefaults() {
    local port
    port="$DM_PORT"
    local issues=""
    [ "$port" = "5236" ] && issues="仍使用默认端口 5236"
    [ "$DM_USER" = "SYSDBA" ] && issues="${issues:+$issues；}仍使用默认管理员账户 SYSDBA"
    if [ -n "$issues" ]; then
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "fail" "$issues" "第1章" "修改默认端口，禁用/加固默认账户。"
    else
        add_result "1.19" "系统安全-数据库" "数据库默认账户/端口" "pass" "未发现默认端口/默认账户问题（端口=${port}，账户=${DM_USER}）。" "第1章" "持续保持非默认端口与加固账户。"
    fi
}

check_1_20_dbpolicy() {
    local pwd_policy
    pwd_policy="$(dm_ini PWD_POLICY)"
    if [ -n "$pwd_policy" ] && [ "$pwd_policy" -ge 15 ] 2>/dev/null; then
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "pass" "已配置口令复杂度策略：PWD_POLICY=${pwd_policy}（>=15，含长度/大小写/数字/特殊字符）。" "第1章" "持续完善口令、审计、加密等安全策略。"
    else
        add_result "1.20" "系统安全-数据库" "数据库安全策略" "fail" "口令复杂度策略不足：PWD_POLICY=${pwd_policy:-未设置}（<15）。" "第1章" "配置口令复杂度策略 PWD_POLICY>=15。"
    fi
}

check_1_21_dbaudit() {
    add_result "1.21" "系统安全-数据库" "数据库操作审计" "manual" "达梦支持审计，行级/列级审计粒度需人工确认审计策略是否覆盖敏感行/列。" "第1章" "配置审计策略覆盖敏感行/列。"
}

check_1_22_monitor() {
    add_result "1.22" "系统安全-数据库" "数据库审计日志留存" "manual" "独立安全监控与审计措施需人工核查部署情况。" "第1章" "部署独立监控/审计平台。"
}

check_1_23_appsep() {
    local listen
    listen="$(ss -lntp 2>/dev/null | awk -v p=":$DM_PORT" '$4 ~ (":"p"$"){print $4}')"
    if [ -z "$listen" ]; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "未检测到监听（可能未安装 ss 或使用自定义端口），请人工确认是否仅为应用服务器提供访问。" "第1章" "限制数据库仅为应用服务器提供访问。"
    elif echo "$listen" | grep -qE '^(0\.0\.0\.0|\*|\[::\])' 2>/dev/null; then
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "manual" "数据库对外暴露监听：$listen，请人工确认是否仅为应用服务器提供访问。" "第1章" "限制数据库仅为应用服务器提供访问。"
    else
        add_result "1.23" "系统安全-数据库" "应用与数据库账户分离" "pass" "数据库未对外暴露监听（$listen），访问受限。" "第1章" "持续确保数据库仅为应用服务器提供访问。"
    fi
}

check_1_24_logaudit180() {
    add_result "1.24" "系统安全-数据库" "日志审计留存180天" "manual" "达梦审计日志留存时长需人工核查是否满足180天（审计日志配置在 dm.ini / 审计归档）。" "第1章" "配置审计日志留存不少于180天。"
}

check_1_25_boundary() {
    local fw fwstate
    fw="$(iptables -S 2>/dev/null | head -20)"
    fwstate="$(firewall-cmd --state 2>/dev/null)"
    if [ -n "$fw" ] || [ "$fwstate" = "running" ]; then
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "manual" "检测到本机防火墙（${fwstate:-iptables 规则存在}），但边界防护/WAF/网络隔离需结合网络架构人工核查。" "第1章" "部署边界防护，限制数据库暴露面。"
    else
        add_result "1.25" "系统安全-数据库" "边界保护抗攻击防篡改" "fail" "未检测到本机防火墙规则，数据库缺乏边界保护。" "第1章" "部署防火墙/WAF 限制暴露面。"
    fi
}

check_1_26_log() {
    local dmhome logdir
    dmhome="${DM_HOME:-}"
    logdir="$(find / -maxdepth 4 -type d -path '*dmdbms*/log' 2>/dev/null | head -1)"
    if [ -n "$logdir" ]; then
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "pass" "检测到达梦日志目录：$logdir。" "第1章" "确保日志完整记录、留存满足要求。"
    else
        add_result "1.26" "系统安全-数据库" "防病毒/补丁日志记录" "manual" "未检测到达梦日志目录（$DM_HOME/log），请人工确认日志记录完整性。" "第1章" "配置达梦日志输出，保留完整日志。"
    fi
}

check_2_16_patchlatest() {
    add_result "2.16" "系统安全-数据库" "补丁修复升级到最新" "manual" "达梦当前版本：${DB_VERSION:-未知}。请人工核对是否已升级到最新版本。" "第2章" "定期升级到最新版本。"
}

# ============================================================
# 主流程
# ============================================================

echo "================================================================"
echo "  配置核查工具 - 达梦数据库(DM8) 版（无需Python）"
echo "  参考标准：配置核查作业指导书v2.0.0"
echo "================================================================"
echo "  连接：$DM_USER@$DM_HOST:$DM_PORT"
echo ""

if [ -z "$DISQL" ]; then
    DM_PRESENT=0
    DB_VERSION="未安装disql"
    echo "  未检测到 disql 客户端，所有项将标记为不适用。"
elif DB_VERSION="$(dm_q "SELECT * FROM V\$VERSION;" 2>/dev/null | grep -oE 'DM Database Server[^[:space:]]* [0-9]+ V[0-9]+' | head -1)"; [ -n "$DB_VERSION" ]; then
    DM_PRESENT=1; CONN_OK=1
    echo "  达梦版本：$DB_VERSION"
else
    DM_PRESENT=1; CONN_OK=0
    echo "  [!] 无法连接达梦数据库（$DM_USER@$DM_HOST:$DM_PORT），所有项将标记为需人工核查。"
    echo "      请通过 DM_USER/DM_PASS 环境变量提供凭据。"
fi
echo ""
echo "开始检查，请稍候..."

if [ "$CONN_OK" = "1" ]; then
    check_1_1_patch
    check_1_7_dbaccounts
    check_1_8_dbprocedures
    check_1_9_dbgrants
    check_1_10_dbaccess
    check_1_11_dbbackup
    check_1_12_dbaudit
    check_1_13_dbisolation
    check_1_15_dbremote
    check_1_16_inputcheck
    check_1_19_dbdefaults
    check_1_20_dbpolicy
    check_1_21_dbaudit
    check_1_22_monitor
    check_1_23_appsep
    check_1_24_logaudit180
    check_1_25_boundary
    check_1_26_log
    check_2_16_patchlatest
else
    db_unreachable "1.1"  "数据库补丁程序"          "升级到最新稳定版本。"
    db_unreachable "1.7"  "数据库账户管理"          "删除冗余账户，设置强口令策略。"
    db_unreachable "1.8"  "数据库存储过程管理"      "删除冗余存储过程。"
    db_unreachable "1.9"  "数据库权限最小化"        "按最小权限原则分配权限。"
    db_unreachable "1.10" "数据库访问控制"          "限制监听地址与访问来源。"
    db_unreachable "1.11" "数据库备份策略"          "建立定期备份。"
    db_unreachable "1.12" "数据库审计插件"          "开启数据库审计。"
    db_unreachable "1.13" "数据库与业务隔离"        "敏感数据分类独立存储。"
    db_unreachable "1.15" "数据库远程访问控制"      "限制超管远程登录。"
    db_unreachable "1.16" "输入验证/SQL注入防护"    "应用层参数化查询。"
    db_unreachable "1.19" "数据库默认账户/端口"     "修改默认端口，加固默认账户。"
    db_unreachable "1.20" "数据库安全策略"          "配置口令/审计/加密策略。"
    db_unreachable "1.21" "数据库操作审计"          "配置行/列级审计。"
    db_unreachable "1.22" "数据库审计日志留存"      "独立监控平台。"
    db_unreachable "1.23" "应用与数据库账户分离"    "限制仅应用服务器访问。"
    db_unreachable "1.24" "日志审计留存180天"       "审计日志留存180天。"
    db_unreachable "1.25" "边界保护抗攻击防篡改"    "部署防火墙。"
    db_unreachable "1.26" "防病毒/补丁日志记录"     "配置日志输出。"
    db_unreachable "2.16" "补丁修复升级到最新"      "升级到最新版本。"
fi

echo ""
echo "================================================================"
print_summary
generate_html
generate_xls
type generate_xlsx >/dev/null 2>&1 && generate_xlsx "$OUT_DIR/配置核查报告_达梦_${STAMP}.xlsx" "达梦数据库配置核查报告"
echo ""
echo "核查完成，请到 output/ 目录查看 HTML/XLS/XLSX 报告。"
