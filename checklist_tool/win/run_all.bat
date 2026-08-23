@echo off
cd /d %~dp0
echo ############################################################
echo #  Windows 配置核查 - 一键运行全部（OS + 数据库 + 中间件）
echo #  VBScript 版，兼容 Windows XP SP3 / 7 / 10（无需 PowerShell）
echo ############################################################

echo.
echo ==================== 1/7 Windows 操作系统核查 ====================
cscript //NoLogo check_xp7.vbs

echo.
echo ==================== 2/7 MySQL 数据库核查 ====================
cscript //NoLogo check_mysql.vbs

echo.
echo ==================== 3/7 Redis 核查 ====================
cscript //NoLogo check_redis.vbs

echo.
echo ==================== 4/7 达梦数据库核查 ====================
cscript //NoLogo check_dm.vbs

echo.
echo ==================== 5/7 SQL Server 核查 ====================
cscript //NoLogo check_sqlserver.vbs

echo.
echo ==================== 6/7 Nginx 中间件核查 ====================
cscript //NoLogo check_nginx.vbs

echo.
echo ==================== 7/7 Tomcat 中间件核查 ====================
cscript //NoLogo check_tomcat.vbs

echo.
echo ==================== 合并生成汇总报告 ====================
cscript //NoLogo merge_report.vbs

echo.
echo ############################################################
echo #  全部核查完成，各组件报告 + 汇总报告在 output 目录下
echo #  汇总：output\配置核查汇总报告_日期时间.xls
echo ############################################################
pause
