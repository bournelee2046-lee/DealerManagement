#!/bin/bash

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 运行同步脚本
./sync_to_github.sh

# 保持窗口打开，让用户看到结果
echo ""
read -p "按回车键退出..."
