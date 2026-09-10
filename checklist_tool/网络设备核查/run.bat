@echo off
rem ============================================================
rem  网络设备核查 - 一键入口（Windows 管理机，系统自带 PowerShell）
rem  用法：run.bat [init / check]
rem    init   第 1 步：生成采集工作区与三厂商命令清单（首次使用）
rem    check  第 2 步：解析设备回显，生成 23 项核查报告
rem    不带参数双击：已采集到回显就直接出报告，否则先生成采集清单
rem  等价入口：check_network.bat（与主包 win\ 同名同源）
rem ============================================================
setlocal
cd /d "%~dp0"

if /i "%~1"=="init"  goto :doinit
if /i "%~1"=="check" goto :docheck
if /i "%~1"=="help"  goto :usage
if /i "%~1"=="/?"    goto :usage

if exist "output\netdev\*.txt" goto :docheck
goto :doinit

:doinit
echo.
echo ############################################################
echo #  网络设备核查 第 1 步：生成采集工作区
echo ############################################################
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_network.ps1" init
if errorlevel 1 goto :psfail
echo.
echo  请在 output\netdev\ 下按命令清单登录各设备采集回显：
echo    - 每台设备存一个 txt，文件名即设备名
echo    - 采集前先在设备上执行 screen-length 0 temporary 取消分页
echo    - 直接粘贴回显，模板中的 CMD 分隔行不要删除
echo  采集完成后再次双击 run.bat 即可输出报告。
echo ############################################################
goto :done

:docheck
echo.
echo ############################################################
echo #  网络设备核查 第 2 步：解析回显并生成报告
echo ############################################################
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check_network.ps1" check
if errorlevel 1 goto :psfail
echo.
echo  报告已输出到 output\ 目录（HTML / XLS，文件名带时间戳）
echo ############################################################
goto :done

:usage
echo.
echo  用法：run.bat [init / check]
echo    init   生成采集工作区与三厂商命令清单（首次使用）
echo    check  解析设备回显并生成核查报告
echo.
goto :done

:psfail
echo.
echo  [错误] PowerShell 未正常执行，请检查：
echo    1) 系统为 Windows 7 及以上，PowerShell 组件未被移除；
echo    2) 本入口已带 -ExecutionPolicy Bypass，若仍被组策略拦截，
echo       请手工执行本目录下同条命令排查。
echo.

:done
echo.
pause
endlocal
