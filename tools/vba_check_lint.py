# -*- coding: utf-8 -*-
"""VBA Common Lib - 离线静态契约检查（不依赖 Excel / win32com）。

把 README「命名规范 / 新增函数 Checklist」转成可执行检查，作为 CI 的快速门禁，
在任何能跑 Python 的机器上验证源码一致性；真正的行为断言仍需 run_tests.py
在真实 Excel 中执行。

规则（对齐 README）：
  R1 模块命名必须为 mod_<域>（VB_Name = "mod_*"）。
  R2 Catalog 双向一致：每个 Public 函数都登记在模块顶部目录注释，
     且注释里的登记项都对应真实定义（防漏登记 / 防幽灵条目）。
  R3 Public 命名规范：函数名前缀必须来自本模块 Catalog 推导的合法前缀集合，
     保证『按域名前缀归位到对应 mod_*』。
  R4 测试覆盖：除测试模块与调试工具模块外，每个 Public 函数都应在
     mod_tests.bas 出现至少一次（TDD：example-as-test）。
     对话框 / 跨工作簿 / 真实文件写入等副作用函数豁免（难以自动化断言）。

用法:
    python tools/vba_check_lint.py --src src [--tests mod_tests.bas]
退出码 = 问题数(ERROR 计 2 / WARN 计 1，按严重度加权)，0 表示全部通过。
"""
from __future__ import annotations

import argparse
import io
import os
import re
import sys

# 副作用 / 交互类函数：真实写文件、弹对话框、跨打开工作簿操作，难以进 TDD 断言。
R4_EXEMPT = {
    "File_Pick": "对话框选文件",
    "Folder_Pick": "对话框选文件夹",
    "File_Copy": "真实写文件",
    "Workbook_AllSheetNames": "跨打开工作簿副作用",
    "Workbook_SaveWithBackup": "真实保存 + 备份文件",
    "Range_JoinedAddress": "多区域，暂无轻量断言",
}

# 模块级豁免 R4 的模块（测试宿主 / 仅输出立即窗口的调试工具）
R4_MODULE_EXEMPT = {"mod_tests", "mod_debug"}

CATALOG_MARKER = "目录 / Catalog"

_PUBLIC_RE = re.compile(r"^\s*Public\s+(?:Function|Sub)\s+([A-Za-z_]\w*)\s*\(", re.M)
_CATALOG_RE = re.compile(r"^\s*'\s+([A-Za-z_]\w*)\s*\(", re.M)
_VBNAME_RE = re.compile(r'^\s*Attribute\s+VB_Name\s*=\s*"([^"]+)"', re.M)
_MOD_PREFIX_RE = re.compile(r"^([A-Za-z]+)_")


def read_utf8(path: str) -> str:
    with io.open(path, "r", encoding="utf-8-sig") as fh:
        return fh.read()


def parse_module(text: str):
    name = None
    m = _VBNAME_RE.search(text)
    if m:
        name = m.group(1)

    # Public 函数
    public_names = set(_PUBLIC_RE.findall(text))

    # Catalog 登记项：从『目录 / Catalog』起到 Option Explicit 前，取带括号的名字
    catalog_names = set()
    cat_idx = text.find(CATALOG_MARKER)
    if cat_idx != -1:
        head = text[:cat_idx]
        tail = text[cat_idx:]
        tail = tail.split("Option Explicit", 1)[0]
        catalog_names = set(_CATALOG_RE.findall(head + tail))
        # 防止目录标记本身上文被误抓：只保留目录段(标记行后的注释)
        catalog_seg = text[cat_idx:].split("Option Explicit", 1)[0]
        catalog_names = set(_CATALOG_RE.findall(catalog_seg))

    return {
        "name": name,
        "public": public_names,
        "catalog": catalog_names,
    }


def check_module(text: str, fname: str, tests_source: str):
    """对单个模块返回 (errors, warnings)。"""
    errors = []
    warnings = []
    info = parse_module(text)
    name = info["name"] or fname

    # R1 模块命名
    if not name or not re.match(r"^mod_[a-z][a-z0-9_]*$", name):
        errors.append(f"[R1] {name or fname}: VB_Name 不符合 mod_<域> 规范")

    public = info["public"]
    catalog = info["catalog"]

    # R2 Catalog 双向一致（仅对使用了目录注释的模块；mod_tests 等测试宿主无此形态）
    if catalog:
        missing = sorted(public - catalog)
        ghost = sorted(catalog - public)
        for fn in missing:
            errors.append(f"[R2] {name}.{fn}: Public 函数未登记进模块目录注释")
        for fn in ghost:
            errors.append(f"[R2] {name}.{fn}: 目录注释登记了但源码中无此函数")

    # R3 Public 命名：前缀须来自本模块 Catalog 推导的合法前缀集合
    if catalog:
        legal = {m for fn in catalog if (m := _MOD_PREFIX_RE.match(fn))}
        for fn in sorted(public):
            m = _MOD_PREFIX_RE.match(fn)
            if m and m.group(1) not in {x.group(1) for x in legal}:
                warnings.append(
                    f"[R3] {name}.{fn}: 前缀 '{m.group(1)}_' 不在本模块合法前缀集合内，"
                    "可能放错了模块"
                )

    # R4 测试覆盖
    if name not in R4_MODULE_EXEMPT:
        for fn in sorted(public):
            if fn in R4_EXEMPT:
                continue
            if fn not in tests_source:
                warnings.append(
                    f"[R4] {name}.{fn}: 未在 mod_tests 找到断言（TDD 要求补一条）"
                )

    return errors, warnings


def main():
    parser = argparse.ArgumentParser(description="VBA 通用库离线静态契约检查")
    parser.add_argument("--src", required=True, help="存放 .bas 的源码目录")
    parser.add_argument("--tests", default=None, help="测试模块文件路径(默认取 src/mod_tests.bas)")
    args = parser.parse_args()

    src = os.path.abspath(args.src)
    if not os.path.isdir(src):
        sys.exit(f"[错误] 源码目录不存在: {src}")

    tests_path = os.path.abspath(args.tests) if args.tests else os.path.join(src, "mod_tests.bas")
    if not os.path.exists(tests_path):
        sys.exit(f"[错误] 测试模块不存在: {tests_path}")
    tests_source = read_utf8(tests_path)

    files = sorted(f for f in os.listdir(src) if f.lower().endswith(".bas"))
    errors = []
    warnings = []
    for fn in files:
        text = read_utf8(os.path.join(src, fn))
        e, w = check_module(text, fn, tests_source)
        errors += e
        warnings += w
        for line in e:
            print(f"[ERROR] {line}")
        for line in w:
            print(f"[WARN ] {line}")

    print(f"----- 静态检查：{len(errors)} 错误 / {len(warnings)} 警告 -----")
    sys.exit(len(errors) * 2 + len(warnings))


if __name__ == "__main__":
    main()