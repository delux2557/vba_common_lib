# -*- coding: utf-8 -*-
"""在真实 Excel 中跑 VBA 单元测试（win32com 驱动）。

流程：全新生成宿主 .xlsm -> 导入 src(/xls) 的 .bas -> Application.Run 调用
指定的测试入口 -> 读取报告文件 -> 打印结果，以失败数作为退出码。

分层说明:
    python tools/run_tests.py --src src --build build
        纯 VBA 层：运行 src/mod_tests.Run_All_Tests
    python tools/run_tests.py --src src --src xls --entry mod_tests_excel --build build
        Excel 绑定层：运行 xls/mod_tests_excel.Run_All_Tests

用法:
    python tools/run_tests.py --src src [--src xls] [--entry mod_tests|mod_tests_excel] --build build
    --src 可重复；--entry 默认取第一个 src 目录里的 mod_tests。
退出码 = 失败用例数（0 表示全部通过）。
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import FORMAT_XLSM, import_bas, quit_excel, start_excel

REPORT_FILE = "vba_tests_report.txt"


def main():
    parser = argparse.ArgumentParser(description="VBA 通用库单元测试(真实 Excel)")
    parser.add_argument("--src", action="append", required=True,
                        help="存放 .bas 的源码目录(可多次传入 src/xls 分层目录)")
    parser.add_argument("--entry", default=None,
                        help="测试入口模块名(默认取首个 src 目录的 mod_tests)")
    parser.add_argument("--build", required=True, help="build 输出目录")
    args = parser.parse_args()

    dirs = [os.path.abspath(d) for d in args.src]
    entry = args.entry or "mod_tests"

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
    print(f"[1/4] 新建宿主: {host_path}")

    # 2. 导入库
    imported = import_bas(wb, dirs)
    print(f"[2/4] 导入模块 {len(imported)} 个")
    wb.Save()

    # 3. 跑测试
    print(f"[3/4] 运行 {entry}.Run_All_Tests ……")
    fail_count = app.Run(f"{entry}.Run_All_Tests", report_path)

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