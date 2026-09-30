#!/bin/bash
# ============================================================
# 配置核查工具 - 网络设备版（华为/华三/锐捷，纯Bash实现）
# 参考标准：配置核查作业指导书v2.0.0（第5章 网络安全 5.1-5.23）
#
# 网络设备为独立硬件，本脚本采用「采集-解析」两步模式：
#   1) bash check_network.sh init    生成采集工作区与三厂商命令清单模板
#      —— 运维陪同登录各设备，按模板粘贴命令回显，每台设备存一个 txt
#   2) bash check_network.sh check   解析回显文件，按 23 项判定，生成报告
#
# 采集文件约定：output/netdev/<设备名>.txt，各命令段以
#   ########## CMD: <命令> ########## 行分隔（模板已带标记，直接粘贴回显即可）
#
# 输出：output/配置核查报告_网络设备_日期时间.html / .xls / .xlsx
# ============================================================
set -u
cd "$(dirname "$0")"

# 零依赖 .xlsx 生成器（可选，需要 zip 命令；缺失时自动降级为 .html/.xls）
[ -f "lib_xlsx.sh" ] && source "lib_xlsx.sh"

OUT_DIR="output"
NET_DIR="$OUT_DIR/netdev"
STAMP="$(date '+%Y%m%d_%H%M%S')"

# ---------- 结果存储 ----------
R_ID=(); R_CAT=(); R_TITLE=(); R_STATUS=(); R_DETAIL=(); R_CHAPTER=(); R_REC=(); R_GUIDE=(); R_METHOD=()
R_COUNT=0

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

add_result() { # id cat title status detail chapter rec
    # 评审整改（2026-09-30 第4/8条）：验证过程/方法列 + 核查项名对齐指导书
    R_COUNT=$((R_COUNT+1))
    R_ID[$R_COUNT]="$1"; R_CAT[$R_COUNT]="$2"
    R_STATUS[$R_COUNT]="$4"; R_CHAPTER[$R_COUNT]="$6"; R_REC[$R_COUNT]="$7"
    R_GUIDE[$R_COUNT]="《配置核查作业指导书v2.0.0》第5章 网络安全 $1"
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

html_esc() { local s="$1"; s="${s//&/&amp;}"; s="${s//</&lt;}"; s="${s//>/&gt;}"; printf '%s' "$s"; }

# ---------- 回显解析辅助 ----------
count_vlans() { # 统计 VLAN 条数：优先数 VLAN 表格行（华为/华三/锐捷 brief 输出），
                # 退化用汇总行 "The total number of vlans is : N"，再退化旧写法 "VLAN N"
    local txt="$1" n
    n="$(echo "$txt" | grep -cE '^[0-9]+[[:space:]]+[^[:space:]]')"
    if [ "${n:-0}" -eq 0 ] 2>/dev/null; then
        n="$(echo "$txt" | grep -ioE 'total number of vlans?[^0-9]*[0-9]+' | grep -oE '[0-9]+' | tail -1)"
    fi
    if [ -z "${n:-}" ]; then
        n="$(echo "$txt" | grep -cE '^(VLAN|vlan) [0-9]+')"
    fi
    printf '%s' "${n:-0}"
}

if_lines() { # 筛出端口状态表里的接口行：兼容全称 GigabitEthernet0/0/1、缩写 GE0/0/1、
             # 华三 XGE1/0/1、锐捷 connected/notconnect 四种写法
    echo "$1" | grep -iE '^[A-Za-z][A-Za-z0-9./:-]*[0-9](/[0-9]+)*[[:space:]]+(up|down|connected|notconnect|dormant|inactive|active|suspended)'
}

# ---------- init：生成采集模板 ----------
do_init() {
    mkdir -p "$NET_DIR"
    cat > "$NET_DIR/采集说明.txt" <<'EOF'
【网络设备配置核查 - 采集操作说明】
1. 每台网络设备（交换机/路由器/防火墙）建一个 txt 文件，文件名即设备名，
   如：核心交换机-01.txt、边界防火墙-01.txt（保存为 UTF-8 或 GBK 均可）。
2. 复制 采集命令清单-华为华三.txt 或 采集命令清单-锐捷.txt 的全部内容到该文件，
   登录设备后逐条执行命令，把每条命令的完整回显粘贴到对应 ########## CMD: 行之后。
   （在设备上可用 terminal monitor / screen-length 0 temporary 之类命令取消分页）
3. 全部设备采集完成后，运行：bash check_network.sh check 生成核查报告。
4. 只采集了部分设备也能出报告，未覆盖项将标记需人工核查。
EOF
    cat > "$NET_DIR/采集命令清单-华为华三.txt" <<'EOF'
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

EOF
    cat > "$NET_DIR/采集命令清单-锐捷.txt" <<'EOF'
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

EOF
    echo "采集工作区已生成：$NET_DIR/"
    echo "  采集说明.txt                    —— 先读这个"
    echo "  采集命令清单-华为华三.txt        —— 华为/华三设备命令模板"
    echo "  采集命令清单-锐捷.txt            —— 锐捷设备命令模板"
    echo "按说明采集完成后运行：bash check_network.sh check"
}

# ---------- check：解析回显并判定 ----------
# 全量回显缓存：所有设备文件拼接，命令段按 ########## CMD: xxx ########## 切分
ALL_TEXT=""
DEV_COUNT=0
DEV_NAMES=""

sec() { # sec "命令关键字" —— 提取全部设备中该命令段（含命令名行之后到下一个CMD标记）
    awk -v cmd="CMD: $1" '
        index($0, cmd) { on=1; buf=""; next }
        /^#+ CMD: / { if(on){ printf "%s\n", buf } on=0 }
        on { buf = buf $0 "\n" }
        END { if(on) printf "%s\n", buf }
    ' <<< "$ALL_TEXT"
}

has_data() { [ -n "$(sec "$1" | tr -d '[:space:]')" ]; }

do_check() {
    if [ ! -d "$NET_DIR" ]; then
        echo "未找到采集目录 $NET_DIR，请先运行：bash check_network.sh init"
        exit 1
    fi
    # 汇总设备回显文件（排除说明与清单）
    for f in "$NET_DIR"/*.txt; do
        [ -e "$f" ] || continue
        base="$(basename "$f")"
        case "$base" in
            采集说明*|采集命令清单*) continue ;;
        esac
        if [ -n "$(grep -c 'CMD:' "$f" 2>/dev/null)" ] && grep -q 'CMD:' "$f" 2>/dev/null; then
            DEV_COUNT=$((DEV_COUNT+1))
            DEV_NAMES="$DEV_NAMES $base"
            ALL_TEXT="$ALL_TEXT
$(cat "$f")"
        fi
    done
    if [ "$DEV_COUNT" -eq 0 ]; then
        echo "采集目录中没有含 CMD: 标记的设备回显文件（当前目录：$NET_DIR）。"
        echo "请按 采集说明.txt 将命令回显粘贴进模板文件后再运行 check。"
        exit 1
    fi
    DEV_NAMES="$(echo "$DEV_NAMES" | sed 's/^ //; s/\.txt//g' | tr '\n' ' ')"
    echo "已加载 $DEV_COUNT 台设备回显：$DEV_NAMES"

    local cfg vlan route ssh_st luser acl dot1x pim mcast brief users
    cfg="$(sec 'display current-configuration'; sec 'show running-config')"
    vlan="$(sec 'display vlan'; sec 'show vlan')"
    route="$(sec 'display ip routing-table'; sec 'show ip route')"
    ssh_st="$(sec 'display ssh server status'; sec 'show ip ssh')"
    acl="$(sec 'display acl all'; sec 'show access-lists')"
    dot1x="$(sec 'display dot1x'; sec 'show dot1x summary')"
    pisol="$(sec 'display port-isolate group')"
    mcast="$(sec 'display multicast routing-table')"
    brief="$(sec 'display interface brief'; sec 'show interfaces status')"
    ike="$(sec 'display ike sa')"
    hrp="$(sec 'display hrp state')"
    ver="$(sec 'display version'; sec 'show version')"

    # ---- 5.1 跨网跨域交换流程（文档类） ----
    add_result "5.1" "网络安全" "跨网跨域数据交换按规定流程进行" "manual" \
        "文档流程类核查：调阅跨网跨域交换审批单与方案评审意见，比对设备侧交换通道与审批范围（见指导书 5.1 方法（4）日志抽查）。" \
        "第5章" "跨网跨域交换须逐次审批并留存方案评审记录。"
    # ---- 5.2 无线组网流程（文档类） ----
    add_result "5.2" "网络安全" "无线网络构建按规定流程执行" "manual" \
        "文档流程类核查：查阅无线建设审批、方案评审，并在授权下探测是否存在未登记私建热点（见指导书 5.2 方法（4））。" \
        "第5章" "无线组网须审批后实施，私建热点应立即清除。"
    # ---- 5.3 最小化架构 ----
    if [ -n "$(echo "$vlan" | tr -d '[:space:]')" ]; then
        local vlan_cnt
        vlan_cnt="$(count_vlans "$vlan")"
        if [ "$vlan_cnt" -le 1 ] 2>/dev/null; then
            add_result "5.3" "网络安全" "按最小化原则设计网络架构" "fail" \
                "设备回显中仅发现 $vlan_cnt 个 VLAN，疑似未按业务划分区域（全设备置于同一广播域）。请核对 topology 与规划表。" \
                "第5章" "按业务划分 VLAN，采用交换技术组网并最小化广播域。"
        else
            add_result "5.3" "网络安全" "按最小化原则设计网络架构" "pass" \
                "回显发现 $vlan_cnt 个 VLAN，已按业务划分。请结合 IP 地址规划表人工确认无整段闲置。" \
                "第5章" "定期复核 VLAN 与 IP 规划的最小化。"
        fi
    else
        add_result "5.3" "网络安全" "按最小化原则设计网络架构" "manual" \
            "未采集到 display vlan / show vlan 回显，无法自动判定。" "第5章" "补采 VLAN 配置后重跑，或按指导书 5.3 人工核查。"
    fi
    # ---- 5.4 边界互联与路由最小化 ----
    if [ -n "$(echo "$route" | tr -d '[:space:]')" ]; then
        local proto_cnt
        proto_cnt="$(echo "$route" | grep -oE 'OSPF|RIP|BGP|ISIS' | sort -u | wc -l)"
        if [ "$proto_cnt" -gt 2 ]; then
            add_result "5.4" "网络安全" "边界互联节点与路由最小化" "fail" \
                "路由表中同时出现 $proto_cnt 种动态路由协议（$(echo "$route" | grep -oE 'OSPF|RIP|BGP|ISIS' | sort -u | tr '\n' '/' )），超出常见业务所需，请逐条核对业务归属。" \
                "第5章" "按业务最小化启用路由协议，关闭无用途协议与闲置网段。"
        else
            add_result "5.4" "网络安全" "边界互联节点与路由最小化" "pass" \
                "动态路由协议共 $proto_cnt 种，未见明显冗余；路由业务归属仍需人工逐条确认。" \
                "第5章" "定期清理无业务归属路由。"
        fi
    else
        add_result "5.4" "网络安全" "边界互联节点与路由最小化" "manual" \
            "未采集到路由表回显，无法自动判定。" "第5章" "补采 display ip routing-table / show ip route。"
    fi
    # ---- 5.5 安全区域划分 ----
    if [ -n "$(echo "$vlan" | tr -d '[:space:]')" ]; then
        local vlan_cnt2
        vlan_cnt2="$(count_vlans "$vlan")"
        if [ "$vlan_cnt2" -ge 3 ] 2>/dev/null; then
            add_result "5.5" "网络安全" "按业务性质划分安全区域" "pass" \
                "发现 $vlan_cnt2 个 VLAN，具备区域划分基础。请比对安全区域划分图确认服务器/存储/嵌入式设备分区。" \
                "第5章" "区域边界应配套隔离控制（ACL/防火墙）。"
        else
            add_result "5.5" "网络安全" "按业务性质划分安全区域" "manual" \
                "VLAN 数量偏少（$vlan_cnt2），是否满足分区要求需结合区域划分图人工确认。" "第5章" "按用途划分安全区域。"
        fi
    else
        add_result "5.5" "网络安全" "按业务性质划分安全区域" "manual" "未采集到 VLAN 回显。" "第5章" "补采 display vlan。"
    fi
    # ---- 5.6 设备安全配置最小化 ----
    if [ -n "$(echo "$cfg" | tr -d '[:space:]')" ]; then
        local telnet_on ftp_on
        telnet_on="$(echo "$cfg" | grep -icE 'telnet server enable|telnet enable')"
        ftp_on="$(echo "$cfg" | grep -icE 'ftp server enable')"
        if [ "$telnet_on" -gt 0 ] || [ "$ftp_on" -gt 0 ]; then
            add_result "5.6" "网络安全" "网络设备最小化安全配置" "fail" \
                "配置中发现 telnet 服务启用 ${telnet_on} 处、ftp 服务启用 ${ftp_on} 处，存在非必要服务；请同时核对账户清单与 ACL 最小放行。" \
                "第5章" "关闭 telnet/ftp 等非必要服务，账户按最小化授权。"
        else
            add_result "5.6" "网络安全" "网络设备最小化安全配置" "pass" \
                "未发现 telnet/ftp 服务启用。账户与 ACL 的最小化仍需按指导书 5.6 方法（4）（5）人工核对。" \
                "第5章" "定期清理冗余账户与全放行策略。"
        fi
    else
        add_result "5.6" "网络安全" "网络设备最小化安全配置" "manual" "未采集到设备配置回显。" "第5章" "补采 display current-configuration / show running-config。"
    fi
    # ---- 5.7 唯一管理服务 ----
    if [ -n "$(echo "$ssh_st" | tr -d '[:space:]')" ]; then
        local telnet_cfg
        telnet_cfg="$(echo "$cfg" | grep -icE 'telnet server enable|telnet enable')"
        if echo "$ssh_st" | grep -qiE 'enable|active|start'; then
            if [ "${telnet_cfg:-0}" -gt 0 ]; then
                add_result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "fail" \
                    "SSH 已启用，但配置中 telnet 仍开启（${telnet_cfg} 处），管理服务不唯一；管理终端 ACL 限制请按 VTY 配置人工核对。" \
                    "第5章" "仅保留 SSH，关闭 Telnet，并以 ACL 限定管理终端。"
            else
                add_result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "pass" \
                    "SSH 管理服务启用且未见 telnet。管理终端来源限制（VTY ACL）需人工核对（见指导书 5.7 方法（3））。" \
                    "第5章" "VTY 线路绑定 ACL 仅放行管理网段。"
            fi
        else
            add_result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "fail" \
                "SSH 服务状态回显未显示启用，且存在明文管理风险。请核实管理方式。" "第5章" "启用 SSH 并关闭明文管理服务。"
        fi
    else
        add_result "5.7" "网络安全" "开放唯一网络管理服务并限制管理终端" "manual" "未采集到 SSH 服务状态回显。" "第5章" "补采 display ssh server status / show ip ssh。"
    fi
    # ---- 5.8 加密管理 ----
    if [ -n "$(echo "$ssh_st" | tr -d '[:space:]')" ]; then
        local v2 telnet_cfg2 http_on
        v2="$(echo "$ssh_st" | grep -icE 'version 2|2\.0|SSH2')"
        telnet_cfg2="$(echo "$cfg" | grep -icE 'telnet server enable|telnet enable')"
        http_on="$(echo "$cfg" | grep -icE 'http server enable|web-manager http')"
        if [ "$v2" -gt 0 ] && [ "${telnet_cfg2:-0}" -eq 0 ] && [ "${http_on:-0}" -eq 0 ]; then
            add_result "5.8" "网络安全" "远程管理采用加密通道" "pass" \
                "SSH 版本为 2.0，未发现 telnet/http 明文管理。防火墙设备 HTTPS 证书请人工核对。" \
                "第5章" "设备与防护设备管理均应走 SSH/HTTPS。"
        else
            add_result "5.8" "网络安全" "远程管理采用加密通道" "fail" \
                "加密管理不达标：SSH2 特征 $v2 处、telnet ${telnet_cfg2:-0} 处、http 管理 ${http_on:-0} 处。存在明文管理风险。" \
                "第5章" "关闭 Telnet/HTTP 管理，统一 SSH2/HTTPS。"
        fi
    else
        add_result "5.8" "网络安全" "远程管理采用加密通道" "manual" "未采集到 SSH 状态回显。" "第5章" "补采 SSH 状态并核对加密层次。"
    fi
    # ---- 5.9 异构防护（台账类） ----
    add_result "5.9" "网络安全" "同功能防护设备异构部署" "manual" \
        "需设备台账比对：同链路承担相同功能的设备应不同品牌/架构。可参考已采集的 display version 型号信息：$(echo "$ver" | grep -oE 'HUAWEI|Huawei|H3C|Comware|Ruijie|RG[- ]?[A-Z0-9]+|S[0-9]{4}|USG[0-9]+|SecPath' | sort -u | tr '\n' ' ')" \
        "第5章" "关键链路同功能设备应异构，防共因失效。"
    # ---- 5.10 组播控制 ----
    local mcast_no_entry
    mcast_no_entry="$(echo "$mcast" | grep -icE 'total 0 entry|0 entry altogether|no (multicast )?route|无(组播)?(表项|路由)|^\(空')"
    if [ -z "$(echo "$mcast" | tr -d '[:space:]')" ] || [ "$mcast_no_entry" -gt 0 ]; then
        add_result "5.10" "网络安全" "组播源/地址/成员控制" "pass" \
            "未发现组播路由表项（无表项或未启用组播业务）。" \
            "第5章" "后续启用组播须先登记并配置边界过滤。"
    else
        add_result "5.10" "网络安全" "组播源/地址/成员控制" "manual" \
            "存在组播路由表项，需按登记清单核对组播源与成员控制（见指导书 5.10 方法（4）边界过滤核查）。" \
            "第5章" "组播边界应配置过滤，未登记组播应清除。"
    fi
    # ---- 5.11 设备备份 ----
    if [ -n "$(echo "$hrp" | tr -d '[:space:]')" ]; then
        add_result "5.11" "网络安全" "重要网络与防护设备备份" "pass" \
            "采集到 HRP（双机热备）状态回显，具备热备机制。冷备设备与配置备份周期仍需按台账人工核对。" \
            "第5章" "备份设备定期启用验证，配置按期备份。"
    else
        add_result "5.11" "网络安全" "重要网络与防护设备备份" "manual" \
            "未采集到双机热备状态（或该设备无热备）。请按指导书 5.11 核对备份制度、冷备与配置备份存档。" \
            "第5章" "重要设备应有备份并演练恢复。"
    fi
    # ---- 5.12 远程传输两层加密 ----
    if [ -n "$(echo "$ike" | tr -d '[:space:]')" ]; then
        add_result "5.12" "网络安全" "远程传输两层加密保护" "manual" \
            "采集到 IKE SA（IPSec 隧道协商）信息，网络层加密在位。链路层/信源层加密层次需按方案人工确认（见指导书 5.12）。" \
            "第5章" "远程传输应满足两层加密（链路/网络/信源任两层）。"
    else
        add_result "5.12" "网络安全" "远程传输两层加密保护" "manual" \
            "未采集到 IKE SA 回显（无 IPSec 隧道或未采集）。按传输加密方案人工核查层次组合。" \
            "第5章" "补采 display ike sa 或核对加密机部署。"
    fi
    # ---- 5.13 管理中心（管理类） ----
    add_result "5.13" "网络安全" "网络安全管理中心与指定管理终端" "manual" \
        "管理类核查：实地查看管理中心与专用终端，登录管理平台核对纳管清单与终端 IP（见指导书 5.13）。" \
        "第5章" "多类设备应经管理中心统一单独管理。"
    # ---- 5.14 等级隔离 ----
    add_result "5.14" "网络安全" "与外部网络按等级物理/逻辑隔离" "manual" \
        "隔离类核查：核对隔离拓扑、实地查看物理隔离点；防火墙会话表（display firewall session table）可用于比对跨域通道，本次$( [ -n "$(sec 'display firewall session table' | tr -d '[:space:]')" ] && echo '已采集到会话表，可比对未审批通道' || echo '未采集到防火墙会话表')。" \
        "第5章" "隔离方式须与防护等级对应。"
    # ---- 5.15 边界防护与告警审计 ----
    add_result "5.15" "网络安全" "边界防护设备的交换控制/告警/审计/阻断能力" "manual" \
        "能力类核查：登录防火墙/网闸查看白名单策略、告警配置与审计日志（见指导书 5.15 方法（2）-（5））。" \
        "第5章" "边界应具备实时告警、审计、阻断与定位能力。"
    # ---- 5.16 细粒度访问控制 ----
    if [ -n "$(echo "$acl" | tr -d '[:space:]')" ]; then
        local acl_rules port_rules
        acl_rules="$(echo "$acl" | grep -cE '^ *(rule|Rule|permit|deny) ')"
        port_rules="$(echo "$acl" | grep -icE 'eq (80|443|22|23|3389|[0-9]+)|destination-port|port (eq|range)')"
        if [ "$acl_rules" -ge 3 ] 2>/dev/null; then
            add_result "5.16" "网络安全" "全网细粒度访问控制" "pass" \
                "ACL 规则共 $acl_rules 条，其中端口级规则特征 $port_rules 处，具备细粒度基础。五元组颗粒度与区域间覆盖需人工抽查（见指导书 5.16 方法（4）（5））。" \
                "第5章" "区域间与主机间均应部署五元组级白名单策略。"
        else
            add_result "5.16" "网络安全" "全网细粒度访问控制" "fail" \
                "ACL 规则仅 $acl_rules 条，颗粒度或覆盖不足，疑似粗放策略或未部署访问控制。" \
                "第5章" "按最小放行原则细化区域间访问控制。"
        fi
    else
        add_result "5.16" "网络安全" "全网细粒度访问控制" "manual" "未采集到 ACL 回显。" "第5章" "补采 display acl all / show access-lists。"
    fi
    # ---- 5.17 租用线路三层加密 ----
    if [ -n "$(echo "$ike" | tr -d '[:space:]')" ]; then
        add_result "5.17" "网络安全" "租用线路三层加密保护" "manual" \
            "采集到 IKE SA，网络层加密在位。链路层与信源层加密、密钥管理记录需人工核查（见指导书 5.17 方法（5））。" \
            "第5章" "租线传输应满足三层加密。"
    else
        add_result "5.17" "网络安全" "租用线路三层加密保护" "manual" \
            "未采集到 IKE SA。租用线路如存在，请按三层加密方案核查。" "第5章" "补采 display ike sa。"
    fi
    # ---- 5.18 接入认证 ----
    if [ -n "$(echo "$dot1x" | tr -d '[:space:]')" ]; then
        if echo "$dot1x" | grep -qiE 'enable|enabled|active|已启用'; then
            add_result "5.18" "网络安全" "802.1x 接入认证" "pass" \
                "802.1x 认证已启用。认证要素（端口+IP+MAC 绑定）需登录 Radius/准入平台人工核对（见指导书 5.18 方法（4））。" \
                "第5章" "认证要素至少含端口、IP、MAC。"
        else
            add_result "5.18" "网络安全" "802.1x 接入认证" "fail" \
                "802.1x 状态回显未显示启用（回显首行：$(echo "$dot1x" | head -1 | cut -c1-40)）。" \
                "第5章" "启用 802.1x 或同强度接入认证。"
        fi
    else
        add_result "5.18" "网络安全" "802.1x 接入认证" "manual" "未采集到 dot1x 回显。" "第5章" "补采 display dot1x / show dot1x summary。"
    fi
    # ---- 5.19 终端逻辑隔离 ----
    if [ -n "$(echo "$pisol" | tr -d '[:space:]')" ]; then
        local pgroups
        pgroups="$(echo "$pisol" | grep -cE 'group|Group')"
        add_result "5.19" "网络安全" "用户计算机间逻辑隔离" "pass" \
            "发现端口隔离组特征 $pgroups 处，接入层已部署终端隔离。主机防火墙状态需在终端侧另行核查。" \
            "第5章" "终端间应默认不可互访，互访经集中服务。"
    else
        add_result "5.19" "网络安全" "用户计算机间逻辑隔离" "manual" \
            "未采集到端口隔离组回显（可能采用 PVLAN 等其他隔离方式）。请按指导书 5.19 方法（4）在终端侧核查防火墙，并做互访测试。" \
            "第5章" "接入层应部署端口隔离或等效措施。"
    fi
    # ---- 5.20 全网行为审计 ----
    add_result "5.20" "网络安全" "全网行为审计且日志留存≥180天" "manual" \
        "平台类核查：登录行为审计/上网管理系统核对覆盖范围、最早可查日志时间与要素完整性（见指导书 5.20 方法（2）-（5））。" \
        "第5章" "审计日志留存不少于 180 天且防删改。"
    # ---- 5.21 攻击实时监视 ----
    add_result "5.21" "网络安全" "攻击/违规行为实时监视告警阻断" "manual" \
        "平台类核查：查看 IDS/IPS/态势感知的策略库、实时告警与处置闭环记录（见指导书 5.21 方法（2）-（5））。" \
        "第5章" "应能实时告警、阻断并定位攻击源。"
    # ---- 5.22 存储管理网隔离 ----
    if [ -n "$(echo "$vlan" | tr -d '[:space:]')" ]; then
        add_result "5.22" "网络安全" "数据存储系统管理网与应用网隔离" "manual" \
            "VLAN 回显可用于核对存储管理口/业务口分属不同 VLAN（见指导书 5.22 方法（3）（4））；请结合存储设备台账与连线照片确认。" \
            "第5章" "存储管理流量与业务流量分网。"
    else
        add_result "5.22" "网络安全" "数据存储系统管理网与应用网隔离" "manual" "未采集到 VLAN 回显。" "第5章" "补采 display vlan 并核对存储端口归属。"
    fi
    # ---- 5.23 设备配备合理必要性 ----
    if [ -n "$(echo "$brief" | tr -d '[:space:]')" ]; then
        local total_if down_if ifl
        ifl="$(if_lines "$brief")"
        total_if="$(printf '%s' "$ifl" | grep -c .)"
        down_if="$(printf '%s' "$ifl" | grep -ciE 'down|notconnect')"
        if [ "$total_if" -gt 0 ] 2>/dev/null; then
            local pct_down=$((down_if * 100 / total_if))
            if [ "$pct_down" -ge 50 ]; then
                add_result "5.23" "网络安全" "网络设备配备合理必要性" "manual" \
                    "端口状态：总 $total_if 口中 $down_if 口 DOWN（${pct_down}%）。闲置比例偏高，请按台账逐台甄别必要冗余与无用途设备。" \
                    "第5章" "清退无用途设备与闲置链路。"
            else
                add_result "5.23" "网络安全" "网络设备配备合理必要性" "pass" \
                    "端口状态：总 $total_if 口中 $down_if 口 DOWN（${pct_down}%），未见大面积闲置。设备必要性仍需台账逐台核对。" \
                    "第5章" "定期清点设备用途。"
            fi
        else
            add_result "5.23" "网络安全" "网络设备配备合理必要性" "manual" "端口状态回显解析不出接口行，请人工核对台账。" "第5章" "定期清点设备用途。"
        fi
    else
        add_result "5.23" "网络安全" "网络设备配备合理必要性" "manual" "未采集到端口状态回显。" "第5章" "补采 display interface brief / show interfaces status。"
    fi

    echo ""
    echo "核查完成，共 $R_COUNT 项（网络设备 $DEV_COUNT 台：$DEV_NAMES）"
    generate_html
    generate_xls
    # 与其余组件脚本一致：另出 .xlsx，供 run_all 的汇总（只并 xlsx）纳入第 5 章
    # 注意 generate_xlsx 内部会 cd 到临时目录，必须传绝对路径
    type generate_xlsx >/dev/null 2>&1 && \
        generate_xlsx "$(pwd)/$OUT_DIR/配置核查报告_网络设备_${STAMP}.xlsx" "网络设备配置核查报告"
}

# ---------- 报告生成（与组件脚本同款模板） ----------
json_esc() {
    local s="$1"
    s="${s//\\/\\\\}"; s="${s//\"/\\\"}"
    s="${s//$'\n'/ }"; s="${s//$'\r'/}"; s="${s//$'\t'/ }"
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

REPORT_TAG="网络设备"
REPORT_TITLE="网络设备配置核查报告"

generate_html() {
    local outfile="$OUT_DIR/配置核查报告_${REPORT_TAG}_${STAMP}.html"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
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
  <h1>${REPORT_TITLE}</h1>
  <div class="sub">参考标准：配置核查作业指导书v2.0.0　|　核查方式：设备命令回显解析（第5章 网络安全）</div>
  <div class="meta">
    <div>设备：$(html_esc "$DEV_COUNT") 台（$(html_esc "$DEV_NAMES")）</div>
    <div>核查时间：$(date '+%Y-%m-%d %H:%M:%S')</div>
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
HTMLHEAD
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
    local outfile="$OUT_DIR/配置核查报告_${REPORT_TAG}_${STAMP}.xls"
    local pass=0 fail=0 manual=0 na=0 i
    for ((i=1; i<=R_COUNT; i++)); do
        case "${R_STATUS[$i]}" in pass) pass=$((pass+1));; fail) fail=$((fail+1));; manual) manual=$((manual+1));; na) na=$((na+1));; esac
    done
    {
        cat <<XLSHEAD
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head><meta charset="UTF-8">
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
<p><b>${REPORT_TITLE}</b>　设备：$(html_esc "$DEV_NAMES")　核查时间：$(date '+%Y-%m-%d %H:%M:%S')</p>
<p>参考标准：配置核查作业指导书v2.0.0 第5章 网络安全</p>
<p>合规：$pass　不合规：$fail　需人工核查：$manual　不适用：$na</p>
XLSHEAD
        echo "<table><tr><th>章节</th><th>编号</th><th>类别</th><th>核查项</th><th>结果</th><th>验证过程/方法</th><th>详情</th><th>修改建议</th><th>安全要求</th><th>参考指导书</th></tr>"
        for ((i=1; i<=R_COUNT; i++)); do
            local scn
            case "${R_STATUS[$i]}" in pass) scn="合规";; fail) scn="不合规";; manual) scn="需人工核查";; *) scn="不适用";; esac
            echo "<tr><td>$(html_esc "${R_CHAPTER[$i]}")</td><td>$(html_esc "${R_ID[$i]}")</td><td>$(html_esc "${R_CAT[$i]}")</td><td>$(html_esc "${R_TITLE[$i]}")</td><td>$scn</td><td>$(html_esc "${R_DETAIL[$i]}")</td><td>$(html_esc "${R_REC[$i]}")</td><td>$(html_esc "${R_GUIDE[$i]}")</td></tr>"
        done
        echo "</table></body></html>"
    } > "$outfile"
    echo "Excel(.xls)报告已生成：$outfile"
}

# ---------- 入口 ----------
case "${1:-}" in
    init)  do_init ;;
    check) do_check ;;
    *) echo "用法：bash check_network.sh [init|check]"; echo "  init  生成采集工作区与三厂商命令清单（首次使用）"; echo "  check 解析已采集的设备回显并生成核查报告" ;;
esac
