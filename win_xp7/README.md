# 配置核查工具 - Windows XP/7 版

适用于**无Python环境**的 Windows XP SP3 / Windows 7 / Windows Server 2003/2008 R2。

使用系统内置的 VBScript（cscript.exe）+ WMI + 注册表实现自动化核查，**无需安装任何依赖**。

---

## 快速开始

**双击运行（推荐）：**

```
run.bat
```

> 建议右键 `run.bat` → 以管理员身份运行，否则部分注册表/WMI查询结果可能不准确。

**命令行运行：**

```cmd
cscript //NoLogo check_xp7.vbs
```

---

## 核查范围

| 章节 | 类别 | 主要检查项 |
|------|------|----------|
| 第1章 | 系统安全 | 补丁状态、防病毒（WMI SecurityCenter）、危险服务、防火墙（XP/7均支持）、RDP加密、Redis端口、审计策略、Defender |
| 第2章 | 用户安全 | 密码到期、账户唯一性、NTFS文件系统、高风险软件（TeamViewer等）、USB管控、口令策略、无线服务、非法外联 |
| 第3章 | 数据安全 | BitLocker加密（Win7）、日志大小、共享检查 |
| 第4章 | 应用安全 | NLA身份验证、三权分立组、IP限制、账户锁定、端口分离 |
| 第5-9章 | 网络/物理/组织/制度/管理 | 标记为手动核查 |

**注意事项：**

- `BitLocker` 检查仅适用于 Windows 7（XP 自动标记为不适用）
- `NLA` 检查仅适用于 Windows 7（XP 自动标记为不适用）
- 防火墙检查：XP 使用 `netsh firewall`，Win7 使用 `netsh advfirewall`，自动适配
- 审计策略通过 `secedit /export` 读取，XP 和 Win7 均兼容

---

## 输出文件

报告保存至 `output\配置核查报告_日期时间.html`，运行完成后自动在浏览器中打开。

---

## 文件说明

```
win_xp7/
├── check_xp7.vbs   核查主脚本（VBScript）
├── run.bat         双击启动入口
└── output/         报告输出目录（自动创建）
```
