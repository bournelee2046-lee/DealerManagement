#!/bin/bash

# 项目目录
PROJECT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$PROJECT_DIR"

echo "========================================="
echo "   异常店治理工具 - GitHub 同步脚本"
echo "========================================="
echo ""

# 检查 git 是否安装
if ! command -v git &> /dev/null; then
    echo "❌ 错误: 未找到 git 命令，请先安装 Git"
    exit 1
fi

# 检查是否为 git 仓库
if [ ! -d ".git" ]; then
    echo "❌ 错误: 当前目录不是 Git 仓库"
    exit 1
fi

# 检查是否配置了 remote
if ! git remote | grep -q "origin"; then
    echo "⚠️  未配置远程仓库"
    echo ""
    read -p "请输入 GitHub 仓库 URL (例如: https://github.com/用户名/仓库名.git): " REPO_URL
    
    if [ -z "$REPO_URL" ]; then
        echo "❌ 错误: 仓库 URL 不能为空"
        exit 1
    fi
    
    git remote add origin "$REPO_URL"
    echo "✅ 已添加远程仓库: $REPO_URL"
    echo ""
fi

# 获取当前分支
BRANCH=$(git branch --show-current)
if [ -z "$BRANCH" ]; then
    BRANCH="main"
fi

echo "📦 当前分支: $BRANCH"
echo ""

# 检查是否有变更
if git diff --quiet && git diff --cached --quiet; then
    echo "ℹ️  没有文件变更需要提交"
else
    echo "📝 发现文件变更，正在添加..."
    git add .
    
    # 获取提交信息
    read -p "请输入提交信息 (默认: 更新代码): " COMMIT_MSG
    if [ -z "$COMMIT_MSG" ]; then
        COMMIT_MSG="更新代码"
    fi
    
    git commit -m "$COMMIT_MSG"
    echo "✅ 已提交变更"
fi

echo ""
echo "🚀 正在推送到 GitHub..."
echo ""

# 推送代码
if git push -u origin "$BRANCH" 2>&1; then
    echo ""
    echo "✅ 同步成功！"
else
    echo ""
    echo "❌ 推送失败"
    echo "提示: 如果是首次推送，可能需要先在 GitHub 创建仓库"
    echo "或者使用: git push -u origin $BRANCH --force"
    exit 1
fi

echo ""
echo "========================================="
echo "   同步完成！"
echo "========================================="
