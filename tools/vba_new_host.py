# -*- coding: utf-8 -*-
"""新建一个空宿主 .xlsm（备用工具；run_tests 内部已内置此逻辑）。

用法:
    python tools/vba_new_host.py --out build/MyHost.xlsm --sheet Sheet1
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import new_host, quit_excel


def main():
    parser = argparse.ArgumentParser(description="新建空宿主 .xlsm")
    parser.add_argument("--out", required=True, help="输出 .xlsm 路径")
    parser.add_argument("--sheet", default="Sheet1", help="首个工作表名")
    args = parser.parse_args()
    try:
        new_host(os.path.abspath(args.out), args.sheet)
        print(f"已新建宿主: {os.path.abspath(args.out)}")
    finally:
        quit_excel()


if __name__ == "__main__":
    main()