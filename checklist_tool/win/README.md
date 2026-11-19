# 配置核查工具 - Windows 版

适用于 **Windows XP SP3 / Windows 7 / Windows 10**，纯 VBScript 实现，**无需 PowerShell、无需 Python**，双击即可运行。

核查对象：Windows 操作系统 + 6 个数据库/中间件（MySQL、Redis、达梦、SQL Server、Nginx、Tomcat）。

---

## 快速开始

**一键核查全部（含汇总）**：以管理员身份双击 `run_all.bat`

> 依次运行：操作系统 → MySQL → Redis → 达梦 → SQL Server → Nginx → Tomcat → 自动合并汇总报告。
>
> `run_all.bat` 会先探测本机 PowerShell 主版本：**3.0 及以上走 `.ps1` 版（报告 10 列）**，否则（XP SP3、未升级 WMF 的 Win7）**回退 `.vbs` 版（报告 8/7 列）**。
> 强制指定：`run_all.bat /ps` 或 `run_all.bat /vbs`。回退时脚本会在首屏明确打印当前模式与检测到的版本号。

**只查操作系统**：以管理员身份双击 `run.bat`

> 建议用管理员权限运行，否则防火墙、账户策略等检查结果可能不准确。

---

## 目录结构

```
win/
├── check_xp7.vbs        Windows 操作系统核查（139 项：136 项编号项 + 3 项补充项）
├── check_mysql.vbs      MySQL / MariaDB 核查（19 项）
├── check_redis.vbs      Redis 核查（16 项）
├── check_dm.vbs         达梦 DM8 核查（19 项）
├── check_sqlserver.vbs  SQL Server 核查（18 项）
├── check_nginx.vbs      Nginx 中间件核查（8 项）
├── check_tomcat.vbs     Tomcat 中间件核查（8 项）
├── merge_report.vbs     汇总合并器（合并各组件报告）
├── db_config.conf       数据库连接配置（改密码看这里）
├── run_all.bat          一键运行全部（含汇总；netdev 有采集回显时自动含网络设备）
├── run.bat              只核查操作系统
├── check_network.ps1    网络设备核查（PowerShell，采集-解析模式，第5章 23 项）
├── check_network.bat    网络设备核查双击入口（init / check）
├── check_*.ps1          PS 版组件脚本（run_all.bat 在 PS 3.0+ 时优先使用，报告 10 列）
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

合并多台被测机的报告：把各机的 `output` 拷成子目录，再合并父目录：

```bat
cscript //NoLogo merge_report.vbs D:\核查汇总
```

（`D:\核查汇总\服务器A\`、`服务器B\` 各放一套组件报告；汇总表首列为「设备」，其余列与组件报告一致：组件/章节/编号/类别/核查项/结果/验证过程·方法/详情/修改建议/核查要求/参考指导书/数据目录（数据库类报告填充，其余为「-」）/报告文件（该行来源报告的文件名）。单目录合并时设备列显示「本机」。）

## 网络设备核查（check_network.ps1 / .bat，Win7+）

交换机、防火墙等网络设备是独立硬件，无法在被测机上直接跑脚本，采用**采集-解析**两步模式（与麒麟版 check_network.sh 同判定逻辑，采集文件两边通用）：

```bat
check_network.bat init     :: 1. 生成采集工作区 output\netdev\（说明 + 华为华三/锐捷命令清单模板）
                           :: 2. 运维陪同登录设备，按清单采集回显粘贴进设备txt（每台设备一个文件）
check_network.bat check    :: 3. 解析回显 → 第5章 23 项判定 → HTML+XLS 报告（带时间戳）
```

23 项判定能力：11 项设备命令回显类自动判定，6 项为脚本采集部分信号或终端侧未覆盖、结论需台账比对（型号/IKE SA/会话表/VLAN 划分、5.19 终端侧未覆盖），6 项审批/平台类需人工核查。设备文件保存为 UTF-8 或 ANSI（记事本默认）均可，脚本自动识别编码。`run_all.bat` 检测到 `output\netdev\` 有采集文件时自动纳入同一轮。

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
3. **报告列数按运行模式**：`run_all.bat` 探测到 PowerShell 3.0+ 时走 `.ps1`，6 个组件报告与 `check_network.ps1` 均为 10 列（含「验证过程/方法」「修改建议」「核查要求」），四类数据库报告另在表头区标注「数据目录」；回退 `.vbs` 时，6 个组件脚本的 HTML 报告为 8 列、XLS 报告为 7 列，`check_xp7.vbs` 仍为 10 列。三个合并器按表头文字自适应 7 / 8 / 10 列输入，统一输出 14 列汇总表（末两列为「数据目录」「报告文件」），因此汇总口径不受运行模式影响。
4. 脚本为 GBK 编码，请用支持 GBK 的编辑器查看，避免用记事本另存为 UTF-8 破坏中文注释。
