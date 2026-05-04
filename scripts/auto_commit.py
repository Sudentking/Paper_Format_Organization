import subprocess
import sys
from datetime import datetime


def git_commit(message=None):
    if not message:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        message = f"auto: 更新代码 {timestamp}"

    commands = [
        ["git", "add", "-A"],
        ["git", "commit", "-m", message],
    ]

    for cmd in commands:
        result = subprocess.run(cmd, capture_output=True, text=True, cwd="d:\\code\\word_脚本\\word-formatter")
        if result.returncode != 0 and "nothing to commit" not in result.stdout:
            print(f"Error: {result.stderr}")
            return False
        if result.stdout:
            print(result.stdout.strip())

    print(f"\n提交成功: {message}")
    return True


if __name__ == "__main__":
    msg = " ".join(sys.argv[1:]) if len(sys.argv) > 1 else None
    git_commit(msg)
