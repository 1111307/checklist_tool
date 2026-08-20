# 配置核查工具

面向 Windows 和国产操作系统（麒麟）的安全配置核查脚本合集，参考《配置核查作业指导书》标准，对**操作系统 + 数据库 + 中间件**进行自动化核查，并生成 HTML / Excel 格式报告。

两个平台的核查项**编号、章节、四态判定规则完全对齐**，方便跨平台核查结果统一汇总对照。

---

## 快速开始

| 平台 | 一键核查全部（OS + 数据库 + 中间件 + 汇总） | 只查操作系统 |
|------|------|------|
| **Windows** | 双击 `win\run_all.bat`（管理员） | 双击 `win\run.bat` |
| **麒麟** | `sudo bash kylin/run_all.sh` | `sudo bash kylin/run.sh` |

> 两个平台均**无需安装 Python 或任何第三方依赖**：Windows 用系统自带的 VBScript（`cscript`），麒麟用系统自带的 Bash。

---

## 核查对象

除操作系统外，另覆盖 6 个数据库/中间件：

| 组件 | Windows 脚本 | 麒麟脚本 | 核查项数 |
|------|------|------|------|
| 操作系统 | `check_xp7.vbs` | `check_kylin.sh` | 137 项 |
| MySQL / MariaDB | `check_mysql.vbs` | `check_mysql.sh` | 19 项 |
| Redis | `check_redis.vbs` | `check_redis.sh` | 16 项 |
| 达梦 DM8 | `check_dm.vbs` | `check_dm.sh` | 19 项 |
| SQL Server | `check_sqlserver.vbs` | `check_sqlserver.sh` | 18 项 |
| Nginx | `check_nginx.vbs` | `check_nginx.sh` | 8 项 |
| Tomcat | `check_tomcat.vbs` | `check_tomcat.sh` | 8 项 |

---

## 目录结构

```
配置核查/
├── win/                         Windows 版（VBScript，无需 PowerShell/Python）
│   ├── check_xp7.vbs            操作系统核查
│   ├── check_mysql.vbs 等       6 个数据库/中间件核查脚本
│   ├── merge_report.vbs         汇总合并器
│   ├── db_config.conf           数据库连接配置（改密码看这里）
│   ├── run_all.bat / run.bat    启动入口
│   ├── README.md                使用说明
│   └── output/                  报告输出目录
│
├── kylin/                       麒麟版（Bash，无需 Python）
│   ├── check_kylin.sh           操作系统核查
│   ├── check_mysql.sh 等        6 个数据库/中间件核查脚本
│   ├── lib_xlsx.sh / merge_xlsx.sh  零依赖 xlsx 生成 + 汇总合并
│   ├── db_config.conf           数据库连接配置
│   ├── run_all.sh / run.sh      启动入口
│   ├── README.md                使用说明
│   └── output/                  报告输出目录
│
├── docker/                      Docker 靶机测试环境（可选）
├── gen_xlsx.py                  Python 版 xlsx 转换器（可选，有 Python 时用）
├── 配置核查表_v2.0.0.xlsx        核查项原始表
├── 配置核查表核查项全览.md        135 项核查项全览（含自动化标注）
├── 配置核查作业指导书_v2.0.0.docx  核查标准依据
└── manual_review/               人工核查补充建议文档
```

---

## 数据库连接配置

各数据库脚本的账号密码通过**各平台目录下的 `db_config.conf`** 配置，编辑即可，无需改脚本：

```ini
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASS=你的密码
```

**连接参数优先级：环境变量 > `db_config.conf` > 脚本默认值**

---

## 核查结果四态

每一项核查结果分为四类，两个平台规则一致：

| 状态 | 含义 |
|------|------|
| 合规（pass） | 检查通过，符合安全要求 |
| 不合规（fail） | 检查未通过，存在安全风险 |
| 需人工核查（manual） | 无法通过脚本自动判定，需现场/文档核查 |
| 不适用（na） | 不适用于当前核查对象（如网络设备层面项、业务应用层面项） |

---

## 核查范围（章节总览）

| 章节 | 类别 | 核查方式 |
|------|------|---------|
| 第1章 | 系统安全 | 自动化核查（补丁、防病毒/EDR、防火墙、SSH/RDP加密、数据库安全等） |
| 第2章 | 用户安全 | 自动化核查（密码策略、账户管理、USB管控、外联控制等） |
| 第3章 | 数据安全 | 自动化核查为主，部分业务层面项标记不适用 |
| 第4章 | 应用安全 | 自动化核查为主，部分业务层面项标记不适用 |
| 第5章 | 网络安全 | 全部标记不适用（网络设备层面，超出单机系统核查范围） |
| 第6-9章 | 物理/组织/制度/管理安全 | 标记需人工核查（需现场/文档核查） |
| 第10章 | 密码与传输安全 | 自动化核查（TLS/加密协议版本） |

---

## 输出报告

运行后在对应目录的 `output/` 下生成：

- 各组件：`配置核查报告_组件名_日期时间.html`（网页）+ `.xls`（表格，WPS/Excel 可直接打开）
- 汇总：`配置核查汇总报告_日期时间.xls` / `.xlsx`（合并所有组件，带"组件"列）

麒麟版在系统有 `zip` 命令时额外生成真 `.xlsx`（零依赖）；Windows 版有 Excel/WPS 时用 COM 生成真 `.xlsx`。

每条结果附带对应《配置核查作业指导书》的章节，便于测试人员对照查阅。

---

## 相关文档

- **`配置核查表_v2.0.0.xlsx`** —— 135 项核查项原始表（10 大类、11 个核查对象）
- **`配置核查表核查项全览.md`** —— 核查项全览，含"✅可自动化 / 🟡部分 / ⚪需人工 / ⏭️跳过"标注
- **`配置核查作业指导书_v2.0.0.docx`** —— 核查标准依据文档
- **`manual_review/`** —— 人工核查补充建议

具体使用方式、各组件详细核查项和配置说明，请查看各平台目录下的 README（`win/README.md`、`kylin/README.md`）。
