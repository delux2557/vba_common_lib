# -*- coding: utf-8 -*-
"""把 src/.bas 打包成 .xlam 加载项，供多个工作簿一处引用。

用法:
    python tools/vba_build_xlam.py --src src --out build/VBA_Common.xlam
产出：加载项文件，宿主工作簿通过 VBE > 工具 > 引用 勾选即可调用 mod_* 函数。
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import build_xlam, quit_excel, read_version, start_excel


def main():
    parser = argparse.ArgumentParser(description="构建 VBA 通用库加载项")
    parser.add_argument("--src", required=True, help="源码 .bas 目录")
    parser.add_argument("--out", required=True, help="输出 .xlam 路径")
    args = parser.parse_args()

    start_excel(visible=False)
    try:
        version = read_version(os.path.abspath(args.src))
        out = build_xlam(os.path.abspath(args.src), os.path.abspath(args.out))
        print(f"已生成加载项: {out}  (mod_version={version})")
    finally:
        quit_excel()


if __name__ == "__main__":
    main()