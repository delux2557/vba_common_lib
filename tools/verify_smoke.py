# -*- coding: utf-8 -*-
"""冒烟验证：从真实 Excel 宿主中按模块限定名调用库函数并打印结果。

用法:
    python tools/verify_smoke.py                       # 自动构建宿主后冒烟
    python tools/verify_smoke.py --host <path.xlsm>    # 对指定已存在宿主冒烟
    python tools/verify_smoke.py --src src --src xls   # 自定义源码目录(默认即 src+xls)

宿主文件不存在时自动新建空宿主并导入源码模块，无需手动准备文件；
宿主已存在时直接打开复用（例如 run_tests 的产物）。
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from vba_excel import import_bas, new_host, quit_excel, start_excel

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_HOST = os.path.join(REPO_ROOT, "build", "smoke_host.xlsm")
DEFAULT_SRC = [os.path.join(REPO_ROOT, "src"), os.path.join(REPO_ROOT, "xls")]


def main():
    parser = argparse.ArgumentParser(description="VBA 通用库冒烟验证(真实 Excel)")
    parser.add_argument("--host", default=DEFAULT_HOST,
                        help="宿主 .xlsm 路径（缺省 %(default)s；不存在时自动新建并导入）")
    parser.add_argument("--src", action="append", default=[],
                        help="源码 .bas 目录(可多次传入；缺省 src+xls)")
    args = parser.parse_args()

    host_path = os.path.abspath(args.host)
    src_dirs = [os.path.abspath(d) for d in (args.src or DEFAULT_SRC)]

    app = start_excel(visible=False)
    try:
        if os.path.exists(host_path):
            print(f"[smoke] 复用宿主: {host_path}")
            wb = app.Workbooks.Open(host_path)
        else:
            print(f"[smoke] 宿主不存在，新建并导入模块: {host_path}")
            wb = new_host(host_path)
            import_bas(wb, src_dirs)
            wb.Save()

        r1 = app.Run("mod_date.Date_Stamp", "date")
        r2 = app.Run("mod_string.String_Trim", "  abc  ")
        r3 = app.Run("mod_regex.Regex_Test", ".hello.", "--helloo--")
        r4 = app.Run("mod_array.Array_Contains", [11, 22, 33], 22)
        r5 = app.Run("mod_ml.Ml_LinearReg_Predict", [1.0, 2.0], 3)
        print(f"mod_date.Date_Stamp(date)   = {r1} (len={len(r1)})")
        print(f"mod_string.String_Trim      = '{r2}'")
        print(f"mod_regex.Regex_Test        = {r3}")
        print(f"mod_array.Array_Contains    = {r4}")
        print(f"mod_ml.Ml_LinearReg_Predict = {r5}")
        wb.Close(False)
        print("[smoke] 全部调用成功")
    finally:
        quit_excel()


if __name__ == "__main__":
    main()
