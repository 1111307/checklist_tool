# 配置核查工具 - 麒麟版（中标麒麟 / 银河麒麟）

适用于**无需 Python**的国产操作系统：

- **中标麒麟 NeoKylin**（基于 CentOS/RHEL，使用 `yum`/`rpm`/`dnf`）
- **银河麒麟 Kylin OS**（基于 Ubuntu/Debian，使用 `apt`/`dpkg`）

纯 Bash + 系统自带命令（`systemctl`/`ss`/`grep`/`awk`/`zip` 等）实现自动化核查，**无需安装任何依赖**，自动识别发行版类型和包管理器。

核查对象：麒麟操作系统 + 6 个数据库/中间件（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat）。

---

## 快速开始

**一键核查全部（含汇总）**：

```bash
sudo bash run_all.sh
```

> 依次运行：麒麟操作系统 → MySQL → Redis → 达梦 → SQL Server → Nginx → Tomcat → 自动合并汇总报告。

**只查操作系统**：

```bash
sudo bash run.sh
```

> 建议 root 或 sudo 运行，否则防火墙/审计/PAM/账户锁定等检查结果可能不准确。

---

## 目录结构

```
kylin/
├── check_kylin.sh       麒麟操作系统核查（137 项）
├── check_mysql.sh       MySQL / MariaDB 核查（19 项）
├── check_redis.sh       Redis 核查（16 项）
├── check_dm.sh          达梦 DM8 核查（19 项）
├── check_sqlserver.sh   SQL Server 核查（18 项）
├── check_nginx.sh       Nginx 中间件核查（8 项）
├── check_tomcat.sh      Tomcat 中间件核查（8 项）
├── lib_xlsx.sh          零依赖 xlsx 生成器（需 zip 命令）
├── merge_xlsx.sh        汇总合并器
├── db_config.conf       数据库连接配置（改密码看这里）
├── run_all.sh           一键运行全部（含汇总）
├── run.sh               只核查操作系统
└── output/              报告输出目录（自动创建）
```

---

## 数据库连接配置

各数据库脚本的账号密码通过 **`db_config.conf`** 配置，编辑该文件即可，无需改脚本：

```ini
# MySQL / MariaDB
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASS=你的密码

# Redis
REDIS_HOST=127.0.0.1
REDIS_PORT=6379
REDIS_PASS=你的密码

# 达梦 DM8
DM_HOST=127.0.0.1
DM_PORT=5236
DM_USER=SYSDBA
DM_PASS=你的密码

# SQL Server
MSSQL_HOST=127.0.0.1
MSSQL_PORT=1433
MSSQL_USER=sa
MSSQL_PASS=你的密码
```

**连接参数优先级：环境变量 > `db_config.conf` > 脚本默认值**

想临时用别的密码时，可用环境变量覆盖，例如：

```bash
MYSQL_PASS=临时密码 sudo bash run_all.sh
```

---

## 单个组件运行

```bash
sudo bash check_mysql.sh
sudo bash check_redis.sh
sudo bash check_dm.sh
sudo bash check_sqlserver.sh
sudo bash check_nginx.sh
sudo bash check_tomcat.sh
```

只合并已生成的报告：

```bash
bash merge_xlsx.sh ./output
```

---

## 核查范围（操作系统部分）

| 章节 | 类别 | 主要检查项 |
|------|------|----------|
| 第1章 | 系统安全 | 补丁状态、防病毒/EDR、高危服务、防火墙(firewalld/ufw/iptables)、SELinux/AppArmor、SSH加密、Redis安全、审计策略(auditd)、资源限制、安全日志、数据库安全 |
| 第2章 | 用户安全 | 密码有效期、账户唯一性、多余服务、高风险软件、USB管控(usbguard/udev)、口令复杂度(PAM)、无线管控、外联控制、用户行为审计 |
| 第3章 | 数据安全 | 磁盘加密(LUKS)、日志文件保护、文件共享(Samba/NFS)、存储分区规划 |
| 第4章 | 应用安全 | 三权分立、管理IP限制、账户锁定策略(pam_faillock)、WAF、网页防篡改(AIDE)、会话超时、SSH默认端口/root登录、国产化适配 |
| 第10章 | 密码与传输安全 | TLS/SSH加密协议版本 |
| 第5章 | 网络安全 | 网络设备层面 23 项，标记"不适用" |
| 第6-9章 | 物理/组织/制度/管理 | 标记"需人工核查" |

**说明：**

- 脚本自动检测发行版（`/etc/os-release`、`/etc/kylin-release`、`/etc/neokylin-release`）与包管理器（apt / yum），无需手动指定。
- 每项核查结果分四类：**合规（pass）/ 不合规（fail）/ 需人工核查（manual）/ 不适用（na）**，与 Windows 版规则一致。
- 核查项编号与 Windows 版 `check_xp7.vbs` 保持一致，共 **137 项**，便于两个平台结果对照汇总。

---

## 报告输出

报告保存在 `output/` 目录：

- 各组件：`配置核查报告_组件名_日期时间.html`（网页）+ `.xls`（表格，WPS/Excel 可直接打开）
- 系统有 `zip` 命令时，额外生成真 `.xlsx`（零依赖，无需 Python/openpyxl）
- 汇总：`配置核查汇总报告_日期时间.xlsx`（合并所有组件）

每条结果附带对应《配置核查作业指导书》章节，方便翻阅指导书核对。

---

## 注意事项

1. **连不上自动降级**：客户端未安装 → 标记"不适用"；已安装但连不上 → 标记"需人工核查"，不会报错中断。
2. **达梦** 需 `disql` 客户端；**SQL Server** 需 `sqlcmd`，并在 `db_config.conf` 填对密码。
3. **真 xlsx 依赖 `zip` 命令**（麒麟几乎必带，缺失时 `yum install -y zip` 或 `apt install -y zip`）；无 zip 时报告仍以 `.html`/`.xls` 输出。
4. 脚本为 UTF-8 编码。
