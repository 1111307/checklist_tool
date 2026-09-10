# 配置核查工具 - 麒麟版（中标麒麟 / 银河麒麟）

适用于**无需 Python**的国产操作系统：

- **中标麒麟 NeoKylin**（基于 CentOS/RHEL，使用 `yum`/`rpm`/`dnf`）
- **银河麒麟 Kylin OS**（基于 Ubuntu/Debian，使用 `apt`/`dpkg`）

纯 Bash + 系统自带命令（`systemctl`/`ss`/`grep`/`awk`/`zip` 等）实现自动化核查，**无需安装任何依赖**，自动识别发行版类型和包管理器。

核查对象：麒麟操作系统 + 6 个数据库/中间件（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat）。

---

## 快速开始

**一键核查全部（自动探测 + 含汇总）**：

```bash
sudo bash run_all.sh
```

> 上机只跑这一个脚本：先**自动探测**本机装了哪些数据库/中间件（探测信号：命令 / 进程 / 监听端口 / 安装目录），只核查探测到的组件，最后合并汇总报告。没装的自动跳过，不会报一堆连接错误。
>
> 探测误判时可用环境变量控制：
> - `FORCE_ALL=1` 强制核查全部组件（探测漏检时用）
> - `SKIP_MYSQL=1 SKIP_REDIS=1 SKIP_DM=1 SKIP_MSSQL=1 SKIP_NGINX=1 SKIP_TOMCAT=1` 单独跳过某组件

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
├── check_network.sh     网络设备核查（第5章 23 项，采集-解析模式）
├── lib_xlsx.sh          零依赖 xlsx 生成器（需 zip 命令）
├── merge_xlsx.sh        汇总合并器
├── db_config.conf       数据库连接配置（改密码看这里）
├── run_all.sh           一键运行全部（含汇总；netdev 有采集回显时自动含网络设备）
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

---

## 网络设备核查（check_network.sh）

交换机、防火墙等网络设备是独立硬件，无法在被测机上直接跑脚本，采用**采集-解析**两步模式：

```bash
# 第一步：初始化采集工作区（生成命令清单模板）
bash check_network.sh init
#   output/netdev/采集说明.txt          操作指引
#   output/netdev/采集命令清单-华为华三.txt   16 条命令模板
#   output/netdev/采集命令清单-锐捷.txt       10 条命令模板

# 第二步：运维陪同登录每台设备（终端先取消分页 screen-length 0 temporary）
#   按清单逐条执行命令，把回显粘贴进「设备名.txt」，每台设备一个文件，放到 output/netdev/

# 第三步：解析回显、生成报告
bash check_network.sh check
```

23 项判定能力（与指导书第 5 章、核查表标注一一对应）：

| 类别 | 项数 | 项 |
|------|------|------|
| 自动判定 | 12 | 5.3-5.8、5.10、5.11、5.16、5.18、5.19、5.23（VLAN/路由/SSH/ACL/802.1x/端口隔离/组播/双机热备/端口闲置率等设备命令回显类） |
| 部分信号 | 5 | 5.9 异构部署、5.12/5.17 传输加密层次、5.14 等级隔离、5.22 存储网分离（脚本提取型号/IKE SA/会话表/VLAN 划分，结论需台账比对） |
| 需人工 | 6 | 5.1、5.2、5.13、5.15、5.20、5.21（审批记录、方案文档、平台界面类） |

报告带时间戳输出到 `output/`（HTML 交互版 + XLS），命令清单与指导书第 5 章方法一一对应，可翻回原文。

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
| 第5章 | 网络安全 | 网络设备层面 23 项，由 `check_network.sh` 单独核查（见下节），OS 脚本中标记"不适用" |
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
