@echo off
setlocal
title 配置核查工具 - 一键运行（操作系统 + 数据库 + 中间件）
cd /d "%~dp0"

rem ============================================================
rem  运行模式选择（评审 2026-09-30 第4/5/9条：报告需含验证过程列）
rem    PS  ：PS 脚本版，报告 10 列（含验证过程/方法、修改建议、安全要求），
rem          需要 PowerShell 3.0 及以上（[PSCustomObject] 是 3.0 才有的类型加速器）
rem    VBS ：VBScript 版，兼容 Windows XP SP3 / 未升级 WMF 的 Win7
rem          组件报告为 HTML 8 列 / XLS 7 列，合并后仍统一为 12 列
rem  强制指定：run_all.bat /ps   或   run_all.bat /vbs
rem ============================================================
set "PSEXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "PSMAJ=0"
if exist "%PSEXE%" for /f "usebackq delims=" %%v in (`"%PSEXE%" -NoProfile -Command "$PSVersionTable.PSVersion.Major" 2^>nul`) do set "PSMAJ=%%v"
set "MODE=VBS"
if %PSMAJ% GEQ 3 set "MODE=PS"
if /i "%~1"=="/ps"  set "MODE=PS"
if /i "%~1"=="/vbs" set "MODE=VBS"

echo ############################################################
echo #  Windows 配置核查 - 一键运行全部（OS + 数据库 + 中间件）
if /i "%MODE%"=="PS" (
  echo #  模式：PS 脚本版（检测到 PS 主版本 %PSMAJ%）  报告 10 列
  echo #        含「验证过程/方法」「修改建议」「安全要求」，数据库报告含「数据目录」
) else (
  echo #  模式：VBScript 版（检测到的 PS 主版本：%PSMAJ%，低于 3.0 或不可用）
  echo #        组件报告 8/7 列，合并后为 12 列；如需 10 列请在 Win7+ 升级 WMF 后重跑
)
echo ############################################################

call :step "1/8" "操作系统核查" "check_xp7"
call :step "2/8" "MySQL 数据库核查" "check_mysql"
call :step "3/8" "Redis 核查" "check_redis"
call :step "4/8" "达梦数据库核查" "check_dm"
call :step "5/8" "SQL Server 核查" "check_sqlserver"
call :step "6/8" "Nginx 中间件核查" "check_nginx"
call :step "7/8" "Tomcat 中间件核查" "check_tomcat"

echo.
echo ==================== 8/8 网络设备核查 ====================
if exist "output\netdev\*.txt" (
  if /i "%MODE%"=="PS" (
    "%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_network.ps1" check
  ) else (
    echo   当前模式无 PS 3.0+，跳过网络设备核查（需 Win7+ 且 PS 可用）。
    echo   可单独运行：check_network.bat check
  )
) else (
  echo   未检测到网络设备采集文件（output\netdev\），跳过。
  echo   如需核查：先运行 check_network.bat init 生成采集清单，运维采集设备回显后重跑本脚本。
)

echo.
echo ==================== 合并生成汇总报告 ====================
cscript //NoLogo merge_report.vbs

echo.
echo ############################################################
echo #  全部核查完成，各组件报告 + 汇总报告在 output 目录下
echo #  汇总：output\配置核查汇总报告_日期时间.xls（12 列）
echo ############################################################
pause
exit /b 0

rem ---------- 单个组件：有同名 .ps1 且模式为 PS 时走 PS，否则走 VBS ----------
:step
echo.
echo ==================== %~1 %~2 ====================
if /i not "%MODE%"=="PS" goto :step_vbs
if not exist "%~dp0%~3.ps1" goto :step_vbs
"%PSEXE%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0%~3.ps1"
goto :eof
:step_vbs
cscript //NoLogo "%~dp0%~3.vbs"
goto :eof
