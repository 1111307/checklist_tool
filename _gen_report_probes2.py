# -*- coding: utf-8 -*-
r"""交叉验证报告 6.7 章节新增 7 张「判定 ↔ 实测」对照图（麒麟靶机实跑）。

   probe_mysql_1_11.png   MySQL 1.11 备份策略（需人工核查 ↔ log_bin=ON、crontab 无备份任务）
   probe_mysql_1_15.png   MySQL 1.15 远程访问控制（合规 ↔ root 仅本机）
   probe_mysql_1_19.png   MySQL 1.19 默认端口（不合规 ↔ port=3306）
   probe_mysql_1_20.png   MySQL 1.20 安全策略（合规 ↔ 严格模式 + secure_file_priv 限定）
   probe_redis_1_11.png   Redis 1.11 备份策略（合规 ↔ 已配置 RDB 快照）
   probe_redis_1_19.png   Redis 1.19 默认端口（不合规 ↔ port=6379）
   probe_redis_1_20.png   Redis 1.20 安全策略（不合规 ↔ maxmemory 未限制）

用法：python _gen_report_probes2.py
（依赖：Docker Desktop 运行中、kylin-target 容器内 mysqld 与 redis-server 已拉起、
  /opt/q111.sh q115 q119 q120 四个查询脚本已就位）
"""
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import _gen_term_figs as T  # noqa: E402

CP = "chcp 65001 >nul && "

JOBS = [
    ("核查项 1.11 判定与实测（MySQL 备份策略）", T.WIN_DIR,
     CP + "docker exec kylin-target sh /opt/q111.sh", "probe_mysql_1_11.png", 30),
    ("核查项 1.15 判定与实测（MySQL 远程访问控制）", T.WIN_DIR,
     CP + "docker exec kylin-target sh /opt/q115.sh", "probe_mysql_1_15.png", 30),
    ("核查项 1.19 判定与实测（MySQL 默认端口）", T.WIN_DIR,
     CP + "docker exec kylin-target sh /opt/q119.sh", "probe_mysql_1_19.png", 30),
    ("核查项 1.20 判定与实测（MySQL 安全策略）", T.WIN_DIR,
     CP + "docker exec kylin-target sh /opt/q120.sh", "probe_mysql_1_20.png", 30),
    ("核查项 1.11 判定与实测（Redis 备份策略）", T.WIN_DIR,
     CP + "docker exec kylin-target redis-cli --raw CONFIG GET save"
          " && echo. && docker exec kylin-target redis-cli --raw CONFIG GET appendonly",
     "probe_redis_1_11.png", 30),
    ("核查项 1.19 判定与实测（Redis 默认端口）", T.WIN_DIR,
     CP + "docker exec kylin-target redis-cli --raw CONFIG GET port",
     "probe_redis_1_19.png", 30),
    ("核查项 1.20 判定与实测（Redis 安全策略）", T.WIN_DIR,
     CP + "docker exec kylin-target redis-cli --raw CONFIG GET maxmemory"
          " && echo. && docker exec kylin-target redis-cli --raw CONFIG GET maxmemory-policy",
     "probe_redis_1_20.png", 30),
]


def main():
    os.makedirs(T.OUT, exist_ok=True)
    for job in JOBS:
        print("=== %s -> %s" % (job[0], job[3]))
        rc = T.capture(*job)
        print("   rc=%s" % rc)
        subprocess.call("taskkill /F /IM cmd.exe", shell=True,
                        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        time.sleep(3)
    print("完成")


if __name__ == "__main__":
    main()
