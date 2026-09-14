# -*- coding: utf-8 -*-
"""一键同步：清空目标 .xlsm 工程中的标准模块，再从 src/ 重新导入。

用法:
    python tools/vba_import.py --src src --target build/VBA_Common_test.xlsm
    python tools/vba_import.py --src src --target build/VBA_Common_test.xlsm --keep mod_importer

说明:
    - 保留 ThisWorkbook 与所有工作表代码页；
    - 其余非文档组件全部删除，再按文件名顺序批量 Import；
    - 需要开启『信任对 VBA 工程对象模型的访问』。
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import import_bas, clear_modules, open_or_create_host, quit_excel, start_excel


def main():
    parser = argparse.ArgumentParser(description="VBA 通用库一键同步(清空+导入)")
    parser.add_argument("--src", required=True, help="存放 .bas 的源码目录")
    parser.add_argument("--target", required=True, help="目标 .xlsm 工作簿")
    parser.add_argument("--keep", nargs="*", default=(), help="保留不删除的组件名")
    args = parser.parse_args()

    app = start_excel(visible=False)
    try:
        wb = open_or_create_host(os.path.abspath(args.target))
        print(f"[1/3] 已打开目标: {wb.Name}")

        removed = clear_modules(wb, keep_names=args.keep)
        print(f"[2/3] 清空模块 {len(removed)} 个: {', '.join(removed) if removed else '(无)'}")

        imported = import_bas(wb, os.path.abspath(args.src))
        print(f"[3/3] 导入模块 {len(imported)} 个: {', '.join(imported)}")

        wb.Save()
        print(f"完成，已保存: {os.path.abspath(args.target)}")
    finally:
        quit_excel()


if __name__ == "__main__":
    main()