#!/bin/bash
cd ~/salary-calc || exit 1

# 删除包含 token 的文件
rm -f git-push.py git-push2.py git-push.sh deploy-git.sh git-push3.sh

# 重新 commit 并 push
git add -A
git commit -m "clean: 移除包含密钥的文件，准备正式部署"
git push -u origin main --force 2>&1

echo "---DONE---"
