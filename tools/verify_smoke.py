# -*- coding: utf-8 -*-
"""验证：从真实 Excel 宿主中按模块限定名调用库函数并打印结果（冒烟验证）。"""
import sys
import win32com.client as win32

HOST = r"G:\Users\Think\AppData\Roaming\TRAE SOLO CN\ModularData\ai-agent\work-mode-projects\6aa6ac4b10f850165e8b42c7\vba_common_lib\build\target_demo.xlsm"

x = win32.gencache.EnsureDispatch("Excel.Application")
x.Visible = False
x.DisplayAlerts = False
try:
    wb = x.Workbooks.Open(HOST)
    check = lambda: 0  # noqa
    r1 = x.Application.Run("mod_date.Date_Stamp", "date")
    r2 = x.Application.Run("mod_string.String_Trim", "  abc  ")
    r3 = x.Application.Run("mod_regex.Regex_Test", ".hello.", "--helloo--")
    r4 = x.Application.Run("mod_array.Array_Contains", [11, 22, 33], 22)
    print(f"mod_date.Date_Stamp(date)   = {r1} (len={len(r1)})")
    print(f"mod_string.String_Trim      = '{r2}'")
    print(f"mod_regex.Regex_Test        = {r3}")
    print(f"mod_array.Array_Contains    = {r4}")
    wb.Close(False)
finally:
    x.Quit()