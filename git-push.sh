#!/bin/bash
set -e

cd ~/salary-calc

# 移除旧的 .git 重新来
rm -rf .git

# 配置 git
git config --global user.name "c940508931-max"
git config --global user.email "c940508931-max@users.noreply.github.com"

# 用正确的 token 格式写入凭证
echo "https://c940508931-max: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@github.com" > ~/.git-credentials
git config --global credential.helper store

# 初始化仓库
git init
git add -A

# 写 .gitignore 排除 build 产物
cat > .gitignore << 'EOF'
__pycache__/
*.pyc
build/
dist/
*.spec
*.egg-info/
.eggs/
uploads/
*.xlsx
EOF

git add -f .gitignore

# 只添加需要的文件
git rm -r --cached build/ dist/ __pycache__/ uploads/ *.spec *.xlsx 2>/dev/null || true
git add -A

git commit -m "init: 薪资计算工具 v1.0"
git branch -M main

# push
git remote add origin https://c940508931-max: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@github.com/c940508931-max/salary-calc.git
git push -u origin main --force

echo ""
echo "✅ PUSH 完成！"
