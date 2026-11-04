@echo off
rem 网络设备核查入口（调用 check_network.ps1，需 Win7+ / PowerShell 2.0+）
rem 用法：check_network.bat init   生成采集工作区与命令清单模板
rem       check_network.bat check  解析设备回显并生成核查报告
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_network.ps1" %1
pause
