#!/bin/bash
# 配置核查工具 - 中标麒麟/银河麒麟版 启动入口
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR" || exit 1

if [ "$(id -u)" != "0" ]; then
    echo "[警告] 当前非root用户，部分检查（防火墙/审计/PAM等）结果可能不准确。"
    echo "建议使用: sudo bash run.sh"
    echo ""
fi

mkdir -p output
bash check_kylin.sh
