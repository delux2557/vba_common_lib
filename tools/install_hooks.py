# 安装 git 钩子到 .git/hooks/。用法：python tools/install_hooks.py
from __future__ import annotations

import os
import shutil
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HOOKS_SRC = os.path.join(ROOT, "tools", "hooks")
HOOKS_DST = os.path.join(ROOT, ".git", "hooks")


def install(name: str) -> bool:
    src = os.path.join(HOOKS_SRC, name)
    dst = os.path.join(HOOKS_DST, name)
    if not os.path.exists(src):
        print(f"跳过缺失钩子: {name}")
        return False
    try:
        shutil.copy2(src, dst)
    except OSError as e:
        print(f"写入失败 {dst}: {e}")
        return False
    try:  # 兼容 WSL/Linux 执行位
        os.chmod(dst, 0o755)
    except OSError:
        pass
    print(f"已安装: {dst}")
    return True


if __name__ == "__main__":
    if not os.path.isdir(HOOKS_DST):
        print(".git/hooks 不存在，请先在仓库内运行。")
        sys.exit(1)
    ok = True
    for f in sorted(os.listdir(HOOKS_SRC)):
        if os.path.isfile(os.path.join(HOOKS_SRC, f)):
            ok = install(f) and ok
    sys.exit(0 if ok else 1)