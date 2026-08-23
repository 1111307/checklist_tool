# 配置核查工具 - Windows 版

适用于 **Windows XP SP3 / Windows 7 / Windows 10**，纯 VBScript 实现，**无需 PowerShell、无需 Python**，双击即可运行。

核查对象：Windows 操作系统 + 6 个数据库/中间件（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat）。

---

## 快速开始

**一键核查全部（含汇总）**：以管理员身份双击 `run_all.bat`

> 依次运行：操作系统 → MySQL → Redis → 达梦 → SQL Server → Nginx → Tomcat → 自动合并汇总报告。

**只查操作系统**：以管理员身份双击 `run.bat`

> 建议用管理员权限运行，否则防火墙、账户策略等检查结果可能不准确。

---

## 目录结构

```
win/
├── check_xp7.vbs        Windows 操作系统核查（137 项）
├── check_mysql.vbs      MySQL / MariaDB 核查（19 项）
├── check_redis.vbs      Redis 核查（16 项）
├── check_dm.vbs         达梦 DM8 核查（19 项）
├── check_sqlserver.vbs  SQL Server 核查（18 项）
├── check_nginx.vbs      Nginx 中间件核查（8 项）
├── check_tomcat.vbs     Tomcat 中间件核查（8 项）
├── merge_report.vbs     汇总合并器（合并各组件报告）
├── db_config.conf       数据库连接配置（改密码看这里）
├── run_all.bat          一键运行全部（含汇总）
├── run.bat              只核查操作系统
├── check_*.ps1          PowerShell 版（可选，Win7+ 装了 PowerShell 才用）
├── lib_xlsx.ps1         PowerShell 版 xlsx 生成器（可选）
├── merge_xlsx.ps1       PowerShell 版汇总合并（可选）
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
MSSQL_HOST=localhost
MSSQL_PORT=1433
MSSQL_USER=sa
MSSQL_PASS=你的密码
```

**连接参数优先级：环境变量 > `db_config.conf` > 脚本默认值**

想临时用别的密码时，可用环境变量覆盖，例如：

```bat
set MYSQL_PASS=临时密码
run_all.bat
```

---

## 单个组件运行

```bat
cscript //NoLogo check_mysql.vbs
cscript //NoLogo check_redis.vbs
cscript //NoLogo check_dm.vbs
cscript //NoLogo check_sqlserver.vbs
cscript //NoLogo check_nginx.vbs
cscript //NoLogo check_tomcat.vbs
```

只合并已生成的报告：

```bat
cscript //NoLogo merge_report.vbs
```

---

## 报告输出

报告保存在 `output\` 目录：

- 各组件：`配置核查报告_组件名_日期时间.html`（网页）+ `.xls`（表格，WPS/Excel 可直接打开）
- 有 Excel/WPS 时，额外用 COM 生成真 `.xlsx`
- 汇总：`配置核查汇总报告_日期时间.xls`（合并所有组件，带"组件"列 + 状态着色）

每条结果含四类状态：**合规 / 不合规 / 需人工核查 / 不适用**，并附带对应《配置核查作业指导书》章节。

---

## 注意事项

1. **连不上自动降级**：客户端未安装 → 标记"不适用"；已安装但连不上（密码错/服务未起）→ 标记"需人工核查"，不会报错中断。
2. **达梦** 需要安装 `disql` 客户端；**SQL Server** 需要安装 `sqlcmd` 并在 `db_config.conf` 填对 `sa` 密码。
3. **旧 PowerShell 脚本**（`check_*.ps1`、`lib_xlsx.ps1`、`merge_xlsx.ps1`）保留供 Win7+ 装了 PowerShell 的环境选用，默认流程走 VBScript。
4. 脚本为 GBK 编码，请用支持 GBK 的编辑器查看，避免用记事本另存为 UTF-8 破坏中文注释。
