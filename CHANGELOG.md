# Changelog

本项目版本变更记录，按 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 惯例维护。
格式为 YYYY-MM-DD 的日志、语义化版本号（Semver）。

## [Unreleased]

### Added
- 纯层新增 7 个函数：`Array_IndexOf` / `Array_Remove` / `Array_RemoveAt`、`String_StartsWith` / `String_EndsWith` / `String_LeftPad`、`Dict_FromArray`（`Dict_To_Array` 之逆）
- 测试增至 106 例（纯层 99 + Excel 层 7）
- 新增 pre-push 静态门禁：契约检查（`vba_check_lint.py`）+ 工具脚本 `py_compile`，由 `tools/install_hooks.py` 安装

## [0.1.0] - 2026-09-18

### Added
- 初始 P0–P1 能力：数组 / 集合 / 字符串 / 正则 / 日期 / 文件 / 字典 / 排序 / JSON 等纯层模块
- Excel 绑定层：区域 `mod_range` / 工作簿 `mod_workbook` / UI `mod_ui`
- 调试与测试：`mod_debug`（立即窗口输出）、`mod_tests`（纯层）、`mod_tests_excel`（Excel 层）
- 工具链：`vba_import.py`（一键导入）、`run_tests.py`（win32com 实机测试）、`vba_check_lint.py`（离线契约检查）、`vba_build_xlam.py`（打包加载项）、`vba_new_host.py` / `verify_smoke.py`

### Changed
- 按宿主依赖拆分双层层级：`src/`（纯 VBA，宿主无关）× `xls/`（Excel 绑定）单向依赖
- 收敛 `mod_date` 为 `Date_Stamp(kind)`，移除误导性命名的 `Date_String` 等单行包装
- 加固 `Workbook_SaveWithBackup` 副作用安全（成败均恢复 `DisplayAlerts`）
- 工具链全线支持多目录：`--src src --src xls`