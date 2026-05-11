import os, subprocess, json

TOKEN = "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"
USER = "c940508931-max"
REPO = "salary-calc"

os.chdir(os.path.expanduser("~/salary-calc"))

# 1. 如果有 .git 重新初始化
subprocess.run(["rm", "-rf", ".git"], capture_output=True)
subprocess.run(["git", "init"], capture_output=True)
subprocess.run(["git", "config", "user.name", USER], capture_output=True)
subprocess.run(["git", "config", "user.email", f"{USER}@users.noreply.github.com"], capture_output=True)

# 2. 先验证 token 权限
r = subprocess.run(
    ["curl", "-s", "-H", f"Authorization: token {TOKEN}",
     f"https://api.github.com/repos/{USER}/{REPO}"],
    capture_output=True, text=True
)
repo_info = json.loads(r.stdout)
print(f"Repo: {repo_info.get('full_name', 'ERROR')}")
print(f"Permissions: {repo_info.get('permissions', 'ERROR')}")
print(f"Default branch: {repo_info.get('default_branch', 'ERROR')}")

# 3. 如果仓库有内容，先 pull
r = subprocess.run(["git", "ls-remote", "--heads",
    f"https://{USER}:{TOKEN}@github.com/{USER}/{REPO}.git"],
    capture_output=True, text=True)
print(f"\nRemote branches: {r.stdout[:200] if r.stdout else '(none)'}")

# 4. 添加、提交、push
subprocess.run(["git", "add", "-A"], capture_output=True)
subprocess.run(["git", "commit", "-m", "init: 薪资计算工具 v1.0"], capture_output=True)
# 忽略 __pycache__ 和 build 文件夹
with open(".gitignore", "w") as f:
    f.write("__pycache__/\n*.pyc\nbuild/\ndist/\n*.spec\nuploads/\n")
subprocess.run(["git", "add", ".gitignore"], capture_output=True)
subprocess.run(["git", "commit", "-m", "add gitignore", "--allow-empty"], capture_output=True)
subprocess.run(["git", "branch", "-M", "main"], capture_output=True)

# 5. Push
r = subprocess.run(
    ["git", "push", "-u",
     f"https://{USER}:{TOKEN}@github.com/{USER}/{REPO}.git",
     "main", "--force"],
    capture_output=True, text=True
)
print(f"\nPush stdout: {r.stdout}")
print(f"Push stderr: {r.stderr}")
print(f"Push returncode: {r.returncode}")
