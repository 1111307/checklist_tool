# -*- coding: utf-8 -*-
"""批量给指导书 v2.0.0 补 Top10（含4.1重做，全部从干净document.xml开始）。
关键修正：每条的下一条标题，从该条正文子节(4.x.2)位置之后查找，避免匹配目录区。
"""
import os, zipfile, re

SRC = r"c:\Users\ryan.xiong\Desktop\配置核查\配置核查作业指导书_v2.0.0.docx"
DST = r"c:\Users\ryan.xiong\Desktop\配置核查\配置核查作业指导书_v2.1草稿.docx"
EXTRACT = r"C:\Users\RYAN~1.XIO\AppData\Local\Temp\docx2"
DOCXML = os.path.join(EXTRACT, "word", "document.xml")

data = open(DOCXML, encoding="utf-8").read()

def para(text, bold=False, heading=False):
    # XML转义：文本里的 <>&'" 必须转义，否则破坏document.xml
    text = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")
    rpr = '<w:rPr><w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/>'
    if bold: rpr += '<w:b/><w:bCs/>'
    rpr += '<w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>'
    indent = '<w:ind w:firstLine="560" w:firstLineChars="200"/>' if not heading else ''
    ppr = '<w:pPr><w:keepNext w:val="0"/><w:keepLines w:val="0"/><w:pageBreakBefore w:val="0"/><w:widowControl w:val="0"/><w:kinsoku/><w:wordWrap/><w:overflowPunct/><w:topLinePunct w:val="0"/><w:autoSpaceDE/><w:autoSpaceDN/><w:bidi w:val="0"/><w:adjustRightInd/><w:snapToGrid/>' + indent + '<w:textAlignment w:val="auto"/>' + rpr + '</w:pPr>'
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'

# (eid, 子节文本, 下一条标题文本, 补充段落列表)
ITEMS = [
("4.1","4.1.2中标麒麟","4.2 应用系统软件应及时安装补丁",[
    para("4.1.3 补充核查方法（应用层与数据库层）", heading=True),
    para("【说明】应用系统软件备份可能涉及三个层面，不可只核查操作系统层。以下补充应用层与数据库层的核查方法，与4.1.1/4.1.2的操作系统层方法配合使用。"),
    para("（1）应用层——核查应用自身备份功能与备份配置：", bold=True),
    para("查看应用管理界面是否具备“备份/恢复”功能模块；核查应用配置文件中备份相关参数（如 backup.path、backup.schedule、dump.target）。"),
    para("Win7/WinXP：Select-String -Path D:\\app\\conf\\* -Pattern \"backup|备份|dump\" -ErrorAction SilentlyContinue"),
    para("中标麒麟/银河麒麟：grep -r \"backup\\|备份\\|dump\" /opt/*/conf/*.conf 2>/dev/null"),
    para("（2）数据库层——核查数据库备份策略与备份文件：", bold=True),
    para("核查数据库是否配置定时备份任务、备份文件是否存在且可恢复。"),
    para("MySQL：核查mysqldump定时任务 crontab -l | grep mysqldump；查binlog是否开启 mysql -e \"SHOW VARIABLES LIKE 'log_bin'\"（用于增量恢复）。"),
    para("SQL Server：SELECT database_name, backup_finish_date, type FROM msdb.dbo.backupset ORDER BY backup_finish_date DESC"),
    para("Oracle：RMAN> list backup summary（核查RMAN备份集）"),
    para("（3）恢复能力验证：选取最近一次备份文件，在测试环境执行恢复操作，确认备份可恢复；核查是否存在恢复脚本（find -name \"*restore*\"）。"),
    para("【判定】应用层、数据库层、操作系统层任一层确认有效备份措施即视为满足；仅操作系统层未发现不能直接判不合规，需结合另外两层人工确认。"),
]),
("3.5","3.5.2中标麒麟","3.6 网络边界",[
    para("3.5.3 补充核查方法（应用层与数据库层日志保护）", heading=True),
    para("【说明】本条标题明确列举“网络、系统、应用和用户行为等日志”，除已核查的系统事件日志/syslog外，还需核查应用日志、数据库审计日志、网络设备日志的保护措施。"),
    para("（1）应用层——应用日志读写控制与保护：", bold=True),
    para("核查应用日志目录权限，仅应用服务账户可写、其他只读。Win7/WinXP：icacls \"D:\\tomcat\\logs\"；中标麒麟/银河麒麟：ls -la /var/log/tomcat/ /var/log/nginx/。"),
    para("（2）数据库层——数据库审计日志保护：", bold=True),
    para("MySQL：SHOW VARIABLES LIKE 'audit%'；Oracle：SELECT * FROM dba_audit_trail WHERE ROWNUM<=5；SQL Server：SELECT * FROM sys.server_audits WHERE is_enabled=1。"),
    para("（3）网络层——网络设备日志传输保护：", bold=True),
    para("登录网络设备核查syslog是否配置加密传输(TLS)、日志服务器访问控制。中标麒麟：grep -r \"tls\\|ssl\\|encrypt\" /etc/rsyslog.conf 2>/dev/null。"),
    para("（4）完整性校验——日志防篡改：核查是否部署日志完整性监控(AIDE/FIM)及日志远程转存防本地篡改。"),
    para("【判定】系统日志、应用日志、数据库审计日志、网络设备日志的保护措施均需核查，任一类日志未采取读写控制/加密/完整性校验即存在风险。"),
]),
("3.11","3.11.2中标麒麟","3.12 检查数据采集",[
    para("3.11.3 补充核查方法（数据库层与应用层权限分级）", heading=True),
    para("【说明】数据访问权限分级除操作系统文件权限外，数据常存储于数据库，需核查数据库表/字段级权限及应用角色权限映射。"),
    para("（1）数据库层——数据库权限分级核查：", bold=True),
    para("MySQL：SHOW GRANTS FOR 'user'@'host';；SQL Server：SELECT * FROM sys.database_permissions；PostgreSQL：\\dp classified_table。确认不同密级数据表/字段仅授权对应用户角色。"),
    para("（2）应用层——应用角色与数据权限映射：", bold=True),
    para("核查应用配置中角色-数据表/字段权限映射（grep -r \"role\\|权限\\|classified\" 应用conf）；通过不同角色用户登录应用测试能否访问不同密级数据。"),
    para("（3）访问审计：核查数据库与应用层访问行为是否启用审计（数据库audit_trail/应用操作日志）。"),
    para("【判定】数据访问权限分级需覆盖OS文件、数据库表/字段、应用角色三层，确认按密级分级且访问行为可审计。"),
]),
("3.14","3.14.2中标麒麟","4.1 应用系统软件",[
    para("3.14.3 补充核查方法（大数据平台层统一管控）", heading=True),
    para("【说明】本条标题专门提及“大数据访问统一管控”，除操作系统层权限外，需核查大数据平台(Hadoop/Hive/Ranger)的访问控制。"),
    para("（1）大数据平台层——HDFS/Hive权限与Ranger策略：", bold=True),
    para("核查HDFS ACL：hdfs dfs -getfacl /data/classified；核查Hive权限：beeline -e \"SHOW GRANT\"；核查Apache Ranger策略：通过Ranger Admin界面查看policy列表；核查Kerberos认证：klist。"),
    para("（2）统一管控核查：确认大数据访问是否必须经统一平台(无直连数据库旁路)，ss -tlnp | grep -E ':50070|:10000|:8020'。"),
    para("（3）数据库审计：核查数据库审计策略是否覆盖大数据查询行为。"),
    para("【判定】大数据访问需通过Ranger/HDFS ACL/Hive权限/Kerberos统一管控，存在绕过统一平台的直连通道即不合规。"),
]),
("4.5","4.5.2中标麒麟","4.6 Web应用系统",[
    para("4.5.3 补充核查方法（网络边界设备与应用层限流）", heading=True),
    para("【说明】防DDoS需多层协同：边界设备(网络层)+应用服务器(应用层)+主机(操作系统层)。4.5.1/4.5.2已核查主机防火墙与内核参数，补充边界设备与应用层。"),
    para("（1）网络边界层——边界DDoS防护设备核查：", bold=True),
    para("登录上游防火墙/流量清洗设备，核查是否配置DDoS检测和清洗策略（SYN Flood/UDP Flood/ICMP Flood检测），查看边界设备DDoS告警和处置日志。"),
    para("（2）应用层——应用服务器限流核查：", bold=True),
    para("Nginx：grep -r \"limit_req\\|limit_conn\\|limit_rate\" /etc/nginx/；Tomcat：grep -r \"maxThreads\\|acceptCount\\|maxConnections\" server.xml；核查WAF限速规则。"),
    para("【判定】防DDoS能力需主机+边界设备+应用层多层协同，仅主机防火墙不足以防御DDoS。"),
]),
("4.8","4.8.2中标麒麟","4.9 Web应用系统应具备",[
    para("4.8.3 补充核查方法（专业防篡改系统与文件完整性）", heading=True),
    para("【说明】网页防篡改需核查专业防篡改系统、文件完整性监控、Web目录权限，而非仅查静态页面发布(属4.7)。"),
    para("（1）专业网页防篡改系统：", bold=True),
    para("核查防篡改软件进程/服务。Win7/WinXP：Get-Process | Where-Object {$_.Name -match \"guard\\|protect\\|tamper\\|防篡改\"}；中标麒麟：systemctl list-units | grep -iE \"guard\\|tamper\\|protect\\|防篡改\"。"),
    para("（2）文件完整性监控(FIM)：核查AIDE/FIM是否覆盖Web目录。grep -E \"/var/www\\|html\" /etc/aide/aide.conf。"),
    para("（3）Web目录权限与只读保护：icacls C:\\inetpub\\wwwroot（仅管理员可写、Web服务账户只读）；mount | grep \"/var/www\"（核查只读挂载）。"),
    para("【判定】部署专业防篡改系统或文件完整性监控覆盖Web目录、且Web目录权限收紧，方为满足。"),
]),
("4.11","4.11.2中标麒麟","4.12 应具备访问应用",[
    para("4.11.3 补充核查方法（应用层RBAC与数据库授权）", heading=True),
    para("【说明】本条位于应用安全章，应核查应用自身角色权限管理，而非仅查操作系统本地用户/组。"),
    para("（1）应用层——应用RBAC核查：", bold=True),
    para("登录应用管理界面查看角色列表和权限分配；核查应用配置中角色-权限映射：Select-String -Path D:\\app\\conf\\* -Pattern \"role\\|permission\\|auth\\|权限\"；grep -r \"role\\|permission\\|auth\\|权限\" /opt/*/conf/*.conf。通过不同角色用户登录应用测试功能权限区分。"),
    para("（2）数据库层——数据库用户授权：", bold=True),
    para("MySQL：SELECT user,host FROM mysql.user; SHOW GRANTS FOR 'appuser'@'%';；SQL Server：SELECT * FROM sys.database_principals WHERE type IN ('S','U')；确认应用使用独立受限数据库账号而非DBA高权限账号。"),
    para("【判定】应用自身实现RBAC(角色-功能/数据权限绑定)、数据库账号权限最小化，方为满足。"),
]),
("4.19","4.19.2中标麒麟","4.20 宜使用自主设计",[
    para("4.19.3 补充核查方法（应用层登录失败处理）", heading=True),
    para("【说明】登录失败处理除操作系统的账户锁定/会话超时外，还需核查应用自身的登录失败处理配置。"),
    para("（1）应用层——应用登录失败处理配置：", bold=True),
    para("核查应用配置中登录失败次数限制/账户锁定时间/CAPTCHA after N failures（grep -r \"loginFail\\|lockout\\|maxAttempts\\|captcha\" 应用conf）；核查应用会话超时配置(session-timeout/timeout参数)。"),
    para("（2）功能测试：通过应用界面连续错误登录，验证是否触发账户锁定/CAPTCHA；验证空闲超时自动退出。"),
    para("【判定】应用层与操作系统层均需具备登录失败处理(锁定/超时/自动退出)，任一层缺失即存在风险。"),
]),
("4.25","4.25.2中标麒麟","4.26 应经过代码级",[
    para("4.25.3 补充核查方法（应用层与数据库层日志审计）", heading=True),
    para("【说明】本条标题写“所有访问行为和管理行为”，除已核查的系统审计(auditd/事件日志)外，还需核查应用访问日志和数据库审计日志及其180天留存。"),
    para("（1）应用层——应用访问日志核查：", bold=True),
    para("核查应用访问日志是否存在且记录访问行为：ls -lh /var/log/nginx/access.log；Get-ChildItem D:\\app\\logs -Filter *access*.log。核查应用日志保留≥180天：grep -r \"rotate\\|maxage\\|180\" /etc/logrotate.d/* | grep -i \"nginx\\|tomcat\\|app\"。"),
    para("（2）数据库层——数据库审计日志核查：", bold=True),
    para("Oracle：SELECT * FROM dba_audit_trail；MySQL：SHOW VARIABLES LIKE 'general_log%'；SQL Server：SELECT * FROM sys.server_audits WHERE is_enabled=1。核查数据库审计日志保留时间。"),
    para("（3）多组合查询能力：核查日志是否支持按时间/事件类型/用户多条件组合查询检索(ausearch/aureport/Get-WinEvent筛选)。"),
    para("【判定】系统审计、应用访问日志、数据库审计日志均需启用且保留≥180天，支持多组合查询，方为满足。"),
]),
("4.29","4.29.2中标麒麟","4.30 检查是否具备统一管理",[
    para("4.29.3 补充核查方法（应用层统一权限管理平台）", heading=True),
    para("【说明】用户访问权限统一管理除操作系统的SSSD/本地用户外，还需核查应用层是否部署统一认证/权限平台(IAM/CAS/OAuth2.0)。"),
    para("（1）应用层——统一权限管理平台核查：", bold=True),
    para("核查是否部署统一认证平台(IAM/CAS/OAuth2.0/SSO)：Get-Service | Where-Object {$_.Name -match \"iam\\|cas\\|oauth\\|sso\\|auth\"}；systemctl list-units | grep -iE \"iam\\|cas\\|oauth\\|sso\\|auth\"。核查应用是否对接统一认证：grep -r \"sso\\|cas\\|oauth\\|auth.server\\|iam\\|统一认证\" /opt/*/conf/*.conf。"),
    para("（2）权限变更集中受控：核查应用权限是否通过统一策略配置(而非分散在各应用)；核查权限变更审批流程和操作日志。"),
    para("【判定】应用接入统一认证/权限平台、权限变更集中受控可追溯，方为满足统一管理要求。"),
]),
("4.32","4.32.2中标麒麟","4.33 检查应用软件是否基于",[
    para("4.32.3 补充核查方法（应用层与数据库层备份恢复）", heading=True),
    para("【说明】本条核查所有应用是否具备备份与恢复功能，除4.32.1/4.32.2的方法外，需核查应用自身备份配置、数据库备份恢复能力及实际恢复验证。"),
    para("（1）应用层——应用备份配置：", bold=True),
    para("核查应用配置文件中backup.path/backup.schedule参数：grep -r \"backup\\|备份\\|dump\" /opt/*/conf/*.conf；查看应用管理界面是否有备份/恢复功能模块。"),
    para("（2）数据库层——数据库备份与恢复：", bold=True),
    para("MySQL：核查mysqldump定时任务 crontab -l | grep mysqldump；查binlog mysql -e \"SHOW VARIABLES LIKE 'log_bin'\"。SQL Server：SELECT * FROM msdb.dbo.backupset ORDER BY backup_finish_date DESC。Oracle：RMAN> list backup summary。"),
    para("（3）恢复能力验证：核查恢复脚本存在：find /opt /root -name \"*restore*\" -o -name \"*recover*\"；选取最近备份在测试环境执行恢复操作确认可恢复。"),
    para("【判定】应用自身备份功能、数据库备份策略、OS定时备份任务任一层有效且经恢复验证可恢复，方为满足。仅存在备份文件未经恢复验证不能直接判合规。"),
]),
]

# 第一步：定位每条正文子节位置
inserts=[]
for eid, sub, nxt_title, paras in ITEMS:
    p_sub = data.find(sub)
    if p_sub < 0:
        print(f"[SKIP] {eid} 子节'{sub}'未找到")
        continue
    # 从子节之后找下一条标题（正文区）
    search_text = ">" + nxt_title
    p_nxt = data.find(search_text, p_sub)
    if p_nxt < 0:
        print(f"[SKIP] {eid} 下条'{nxt_title}'在子节后未找到")
        continue
    p_start = data.rfind("<w:p ", 0, p_nxt)
    if p_start < 0: p_start = data.rfind("<w:p>", 0, p_nxt)
    insertion = "".join(paras)
    inserts.append((p_start, insertion, eid))
    print(f"[OK] {eid} 子节{p_sub} 下条{p_nxt} 插入点{p_start} 段落{len(paras)}")

# 倒序插入
inserts.sort(key=lambda x:-x[0])
for p_start, insertion, eid in inserts:
    data = data[:p_start] + insertion + data[p_start:]

open(DOCXML,"w",encoding="utf-8").write(data)
print(f"\n已插入{len(inserts)}条")

# 重新打包（从SRC原始zip复制结构，只替换document.xml）
with zipfile.ZipFile(SRC,"r") as zin:
    names=zin.namelist()
with zipfile.ZipFile(DST,"w",zipfile.ZIP_DEFLATED) as zout:
    for name in names:
        if name=="word/document.xml":
            zout.write(DOCXML, name)
        else:
            with zipfile.ZipFile(SRC,"r") as zin:
                zout.writestr(name, zin.read(name))
print(f"已生成: {DST}")
print(f"大小: {os.path.getsize(DST)//1024} KB")
