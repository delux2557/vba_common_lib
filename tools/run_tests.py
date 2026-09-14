# -*- coding: utf-8 -*-
"""在真实 Excel 中跑 mod_tests 单元测试（win32com 驱动）。

流程：全新生成宿主 .xlsm -> 导入 src/.bas -> Application.Run 调用
mod_tests.Run_All_Tests -> 读取报告文件 -> 打印结果，以失败数作为退出码。

用法:
    python tools/run_tests.py --src src --build build
退出码 = 失败用例数（0 表示全部通过）。
"""
from __future__ import annotations

import argparse
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import FORMAT_XLSM, import_bas, quit_excel, start_excel

REPORT_FILE = "vba_tests_report.txt"


def main():
    parser = argparse.ArgumentParser(description="VBA 通用库单元测试(真实 Excel)")
    parser.add_argument("--src", required=True, help="存放 .bas 的源码目录")
    parser.add_argument("--build", required=True, help="build 输出目录")
    args = parser.parse_args()

    host_dir = os.path.abspath(args.build)
    os.makedirs(host_dir, exist_ok=True)
    host_path = os.path.join(host_dir, "VBA_Common_test.xlsm")
    report_path = os.path.join(host_dir, REPORT_FILE)
    if os.path.exists(host_path):
        os.remove(host_path)
    if os.path.exists(report_path):
        os.remove(report_path)

    app = start_excel(visible=False)

    # 1. 新建空宿主
    wb = app.Workbooks.Add()
    wb.Worksheets(1).Name = "Sheet1"
    wb.SaveAs(host_path, FORMAT_XLSM)
    print(f"[1/3] 新建宿主: {host_path}")

    # 2. 导入库
    imported = import_bas(wb, os.path.abspath(args.src))
    print(f"[2/3] 导入模块 {len(imported)} 个")
    wb.Save()

    # 3. 跑测试
    print("[3/3] 运行单元测试……")
    fail_count = app.Run("mod_tests.Run_All_Tests", report_path)

    wb.Save()
    wb.Close(SaveChanges=False)

    if os.path.exists(report_path):
        # mod_file.File_Write 以 FSO(unicode=True) 写出，即 UTF-16
        with open(report_path, encoding="utf-16") as fh:
            print("\n" + fh.read())
    else:
        print("\n[warn] 未生成测试报告文件（请检查是否开启『信任对 VBA 工程对象模型的访问』）")

    quit_excel()
    sys.exit(0 if fail_count == 0 else fail_count)


if __name__ == "__main__":
    main()