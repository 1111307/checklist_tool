#!/bin/bash
# ============================================================
# 麒麟（Linux）配置核查 - 一键运行全部
# 覆盖：麒麟操作系统 + Linux 数据库 + Linux 中间件
# 用法：bash run_all.sh
# 数据库/中间件连接参数可用环境变量覆盖，例如：
#   MYSQL_HOST=127.0.0.1 MYSQL_PASS=xxx \
#   REDIS_PASS=xxx \
#   DM_USER=SYSDBA DM_PASS=xxx \
#   MSSQL_USER=sa MSSQL_PASS=xxx \
#   bash run_all.sh
# ============================================================
set -u
cd "$(dirname "$0")"

echo "############################################################"
echo "#  麒麟配置核查 - 一键运行全部（OS + 数据库 + 中间件）"
echo "############################################################"

echo ""
echo "==================== 1/7 麒麟操作系统核查 ===================="
bash check_kylin.sh

echo ""
echo "==================== 2/7 MySQL 数据库核查 ===================="
bash check_mysql.sh

echo ""
echo "==================== 3/7 Redis 核查 ===================="
bash check_redis.sh

echo ""
echo "==================== 4/7 达梦数据库核查 ===================="
bash check_dm.sh

echo ""
echo "==================== 5/7 SQL Server 核查 ===================="
bash check_sqlserver.sh

echo ""
echo "==================== 6/7 Nginx 中间件核查 ===================="
bash check_nginx.sh

echo ""
echo "==================== 7/7 Tomcat 中间件核查 ===================="
bash check_tomcat.sh

echo ""
echo "==================== 合并生成汇总 Excel(.xlsx) ===================="
bash merge_xlsx.sh ./output

echo ""
echo "############################################################"
echo "#  全部核查完成，报告在各脚本 output/ 目录下"
echo "#  Excel(.xlsx) 汇总报告位于项目根目录：配置核查汇总报告_*.xlsx"
echo "############################################################"
