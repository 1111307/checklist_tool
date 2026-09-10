# 配置核查工具包（checklist_tool）

面向 Windows 与国产操作系统（麒麟）的安全配置核查工具，配套指导书、核查表与交叉验证报告。两个平台的核查项**编号、章节、四态判定规则完全对齐**，核查结果可跨平台汇总对照。

---

## 快速开始

| 平台 | 一键核查全部（OS + 数据库 + 中间件 + 汇总） | 只查操作系统 |
|------|------|------|
| **Windows**（XP/7/10） | 管理员双击 `win\run_all.bat` | 双击 `win\run.bat` |
| **麒麟**（中标/银河） | `sudo bash kylin/run_all.sh` | `sudo bash kylin/run.sh` |

> **零依赖**：Windows 用系统自带 VBScript（cscript），麒麟用系统自带 Bash，被测机无需安装任何软件。
> 麒麟版 `run_all.sh` 会**自动探测**本机安装的数据库/中间件，只核查存在的组件，没装的自动跳过（`FORCE_ALL=1` 可强制全查）；检测到 `kylin/output/netdev/` 有设备采集回显时，网络设备核查自动纳入同一轮。

数据库/中间件连接参数：编辑各平台目录下的 `db_config.conf`，或用环境变量覆盖（`MYSQL_PASS=xxx sudo bash kylin/run_all.sh`）。

## 核查对象

操作系统之外，覆盖 6 个数据库/中间件与网络设备，共 16 个核查脚本：

| 组件 | Windows 脚本 | 麒麟脚本 |
|------|------|------|
| 操作系统（137 项） | `win\check_xp7.vbs` | `kylin/check_kylin.sh` |
| MySQL / MariaDB | `win\check_mysql.vbs` | `kylin/check_mysql.sh` |
| Redis | `win\check_redis.vbs` | `kylin/check_redis.sh` |
| 达梦 DM8 | `win\check_dm.vbs` | `kylin/check_dm.sh` |
| SQL Server | `win\check_sqlserver.vbs` | `kylin/check_sqlserver.sh` |
| Nginx | `win\check_nginx.vbs` | `kylin/check_nginx.sh` |
| Tomcat | `win\check_tomcat.vbs` | `kylin/check_tomcat.sh` |
| 网络设备（第5章 23 项） | `win\check_network.ps1` | `kylin/check_network.sh` |

### 网络设备核查（采集-解析模式）

交换机、防火墙是独立硬件，无法在被测机上跑脚本，`check_network.sh` 采用三步工作流：

```bash
bash kylin/check_network.sh init     # 1. 生成三厂商（华为/华三/锐捷）命令采集清单
                                      # 2. 运维陪同登录设备，按清单采集回显粘贴进设备txt
bash kylin/check_network.sh check    # 3. 解析回显 → 23 项判定 → HTML+XLS 报告（带时间戳）
```

Windows 管理机同理：`win\check_network.bat init` / `check`（PowerShell 版，判定逻辑与麒麟版一致，采集文件两边通用）。

其中 11 项设备命令回显类自动判定（含端口闲置率），6 项为脚本采集部分信号或部分对象未检测、结论需台账比对（型号/加密层次/会话表/VLAN 划分、5.19 终端侧未覆盖），6 项审批/平台类需人工核查（详见 `kylin\README.md`）。采集回显放入 `kylin/output/netdev/` 后重跑 `run_all.sh`，网络设备核查会自动纳入一键流程。

### 脚本跑不了的机器：人工核查台

`manual_check.html`（单文件、零依赖、浏览器直接打开）：136 项逐项勾选（四态）、章级一键批量、**截图 Ctrl+V 粘贴取证**、草稿自动保存，导出与自动脚本同款格式的 HTML/Excel 带时间戳报告。现场核查脚本无法覆盖的机器（只读系统、组策略拦截等）用它。

## 核查结果四态

| 状态 | 含义 |
|------|------|
| 合规（pass） | 检查通过，符合安全要求 |
| 不合规（fail） | 检查未通过，存在安全风险 |
| 需人工核查（manual） | 无法脚本自动判定，需现场/文档核查 |
| 不适用（na） | 不适用于当前核查对象 |

每条结果附带对应《配置核查作业指导书》的章节出处，便于对照原文。

## 输出报告

运行后在各平台 `output\` 目录生成（首次运行自动创建）：

- **HTML 报告**（推荐）：四状态指标卡、实时搜索、状态筛选、彩色徽章表格，点击卡片即可筛选，支持直接打印归档
- **XLS 报告**：Excel/WPS 直接打开，按"未通过 → 需人工核查 → 不适用 → 通过"四区展示
- **汇总报告**：`run_all` 一键运行后自动合并所有组件结果，带"组件"列

## 目录结构

```
checklist_tool/
├── win\                      Windows 版（VBScript，无需 PowerShell/Python）
│   ├── check_*.vbs/.ps1      核查脚本
│   ├── run.bat / run_all.bat 启动入口
│   └── README.md             详细使用说明
├── kylin\                    麒麟版（纯 Bash，无需 Python）
│   ├── check_*.sh            核查脚本（含网络设备 check_network.sh）
│   ├── run.sh / run_all.sh   启动入口（自动探测组件）
│   └── README.md             详细使用说明（含网络设备采集-解析流程）
├── manual_check.html                       人工核查台（136 项，截图取证 + 双格式导出）
├── 配置核查作业指导书_v2.2.docx             核查标准依据（10 章、136 个编号检查项 + 第5章 13 张终端操作截图）
├── 配置核查表_v2.0.0.xlsx                  检查项 × 11 类核查对象矩阵（含华为交换机等网络设备；√适用/—不适用；第 14 列逐项标注用哪个脚本、哪些对象可自动、哪些需人工）
└── 指导书与核查工具交叉验证报告.docx         指导书与脚本逐项比对（136 项，以脚本是否实际执行核查为准，完全自动化率 31.6%、含部分自动化 69.1%）
```

## 相关文档

- 《配置核查作业指导书》v2.2 —— 核查标准与操作方法（每节含判定标准与双平台核查步骤；与脚本、人工核查台、交叉验证报告标注的参考标准 v2.2 为同一版本，含第5章终端操作截图）
- 《配置核查表》—— 检查项原始矩阵（11 列对象，含华为交换机），标注版第 14 列明确标注每项「可自动化 / 部分可自动化 / 需人工」，含网络设备脚本能力
- 《指导书与核查工具交叉验证报告》—— 可自动化 43 项 / 部分可自动化 51 项 / 需人工 42 项（以「脚本是否实际执行核查」为标准，结合勾选矩阵逐对象核对；表格第 14 列同口径标注，含脚本名与可自动/需人工边界）

具体使用方式、各组件详细核查项和数据库连接配置，见 `win\README.md` 与 `kylin\README.md`。
