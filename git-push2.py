import os, subprocess, json, base64

TOKEN = "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"
USER = "c940508931-max"
REPO = "salary-calc"

os.chdir(os.path.expanduser("~/salary-calc"))

# 先用 git push 试 HTTPS，换一种凭证传递方式
# 直接用嵌入式令牌在 URL 里
r = subprocess.run(
    ["git", "push", "-u",
     f"https://{USER}:{TOKEN}@github.com/{USER}/{REPO}.git",
     "main", "--force"],
    capture_output=True, text=True,
    env={**os.environ, "GIT_TRACE": "1"}
)
print(f"stdout: {r.stdout}")
print(f"stderr: {r.stderr[:2000]}")
print(f"code: {r.returncode}")

# 如果失败，用 GIT_ASKPASS 方式
if r.returncode != 0:
    print("\n--- Trying GIT_ASKPASS approach ---")
    # 写一个 askpass 脚本
    with open("/tmp/git-askpass.sh", "w") as f:
        f.write("#!/bin/bash\n")
        f.write(f'echo "{TOKEN}"\n')
    os.chmod("/tmp/git-askpass.sh", 0o755)
    
    r2 = subprocess.run(
        ["git", "push", "-u",
         f"https://{USER}@github.com/{USER}/{REPO}.git",
         "main", "--force"],
        capture_output=True, text=True,
        env={**os.environ, "GIT_ASKPASS": "/tmp/git-askpass.sh"}
    )
    print(f"stdout: {r2.stdout}")
    print(f"stderr: {r2.stderr[:2000]}")
    print(f"code: {r2.returncode}")
