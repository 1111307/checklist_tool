# 配置核查工具 - 中标麒麟/银河麒麟版

适用于**无需Python**的国产操作系统：

- **中标麒麟 NeoKylin**（基于CentOS/RHEL，使用 `yum`/`rpm`/`dnf`）
- **银河麒麟 Kylin OS**（基于Ubuntu/Debian，使用 `apt`/`dpkg`）

脚本使用纯 Bash + 系统自带命令（`systemctl`/`ss`/`grep`/`awk`等）实现自动化核查，**无需安装任何依赖**，会自动识别发行版类型和包管理器。

---

## 快速开始

```bash
sudo bash run.sh
```

> 建议使用 root 或 sudo 运行，否则防火墙/审计/PAM/账户锁定等配置的核查结果可能不准确。

也可以直接运行主脚本：

```bash
sudo bash check_kylin.sh
```

---

## 核查范围

| 章节 | 类别 | 主要检查项 |
|------|------|----------|
| 第1章 | 系统安全 | 补丁状态、防病毒/EDR、高危服务、防火墙(firewalld/ufw/iptables)、SELinux/AppArmor、SSH加密、Redis安全、审计策略(auditd)、资源限制、安全日志、数据库安全(MySQL/达梦/Kingbase/PostgreSQL) |
| 第2章 | 用户安全 | 密码有效期、账户唯一性、多余服务、高风险软件、USB管控(usbguard/udev)、口令复杂度(PAM)、无线管控、外联控制、用户行为审计 |
| 第3章 | 数据安全 | 磁盘加密(LUKS)、日志文件保护、文件共享(Samba/NFS)、存储分区规划 |
| 第4章 | 应用安全 | 三权分立、管理IP限制、账户锁定策略(pam_faillock)、WAF、网页防篡改(AIDE)、会话超时、SSH默认端口/root登录、国产化适配 |
| 第10章 | 密码与传输安全 | TLS/SSH加密协议版本 |
| 第5章 | 网络安全 | 网络拓扑/边界防护/VLAN/VPN/802.1x等23项，均为网络设备层面，标记"不适用"（单机系统核查范围之外） |
| 第6-9章 | 物理/组织/制度/管理 | 机房安全、组织架构、管理制度、应急演练等，标记"需人工核查"（需现场/文档核查） |

**说明：**

- 脚本会自动检测发行版（读取 `/etc/os-release`、`/etc/kylin-release`、`/etc/neokylin-release`）与包管理器（apt / yum），无需手动指定。
- 数据库相关检查（1.7-1.23）：脚本不内置数据库凭据，会自动探测 MySQL/MariaDB、达梦(DM)、人大金仓(Kingbase)、PostgreSQL 服务是否运行；未运行则标记"不适用"，运行中但需要登录凭据的项标记"需人工核查"。
- 强制访问控制：优先检测 SELinux（`getenforce`），无 SELinux 时检测 AppArmor（`aa-status`），二者皆无则标记人工核查。
- 每项核查结果分为四类：**合规（pass）/ 不合规（fail）/ 需人工核查（manual）/ 不适用（na）**，与 Windows 版规则一致。
- 核查项编号（如 1.1、4.32、5.1-5.23、6.1-6.9 等）与 Windows 版 `check_xp7.vbs` 保持一致，共 **137 项**：约60项为第1-4、10章的技术自动化核查，约77项为第2-10章业务/网络/物理/组织/制度/管理层面的人工核查或不适用项，便于两个平台的核查结果对照汇总。

---

## 输出文件

报告保存至 `output/配置核查报告_日期时间.html` 和 `.xls`（HTML表格格式，Excel/WPS均可直接打开，无需安装Office）。

---

## 文件说明

```
kylin/
├── check_kylin.sh   核查主脚本（Bash）
├── run.sh           启动入口（含root权限提示）
└── output/          报告输出目录（自动创建）
```
