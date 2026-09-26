# Changelog

本项目版本变更记录，按 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 惯例维护。
格式为 YYYY-MM-DD 的日志、语义化版本号（Semver）。

> **版本号唯一来源是 `src/VERSION`**（构建时自动生成 `mod_version` 注入运行期）。
> CHANGELOG 仅在发版时登记版本号与变更内容，不承担"版本定义"职责，防止多来源漂移。

## [Unreleased]

### Added
- `mod_debug` 新增通用对象打印 `Repr(obj[, max_depth])` / `Print_Repr`：Python 风格 repr，
  解决 `Debug.Print` 直接打印对象报错的问题。数组(一维/二维逐元素展开，三维及以上输出
  结构摘要)、集合、字典、Range(地址+值)、普通对象(有默认成员按默认值展示，否则
  `<类型 at 0x地址>`)；标量字符串加引号并转义内部引号，深度默认 16 层防循环引用无限递归
- `mod_range` 新增批量单元格读写：`Range_ReadValues(rng)`（一次读整块，原生 Value 语义：
  单格→标量、整行/整列→一维、多格→二维 1 基数组）、`Range_WriteValues(rng, values[, as_column])`
  （一次写整块：二维数组按尺寸自动 Resize 以左上角为锚点、标量填充全区、一维数组可转置成列；
  写入期间临时关 ScreenUpdating 防闪烁并成败均恢复——替代逐格循环，性能差一个数量级；
  多区域 / Nothing / 空数组 / 对象等非法输入抛错 #45000）
- `mod_date` 新增 UTC 与格式化：`Date_UtcNow`（GetSystemTime 取 UTC，无 COM 依赖）、
  `Date_ToUtc` / `Date_FromUtc`（按当前时区偏移换算，往返可还原）、`Date_StampUtc`（UTC 时间戳，
  与 `Date_Stamp` 同款 kind 格式）、`Date_Format`（统一默认 `yyyy-mm-dd hh:nn:ss` 的 `Format$` 封装）
- `mod_string` 新增模板格式化 `String_Format(template, args...)`：`{0}` 占位符替换、`{{` / `}}` 转义；
  参数越界 / 占位符语法错误 / 多余 `}` 抛错 #45000（调用方错误尽早暴露）

### Fixed
- `Date_ToUtc` / `Date_FromUtc` 参数名避开 VBA 保留字 `local`：保留字作标识符会使该过程无法编译、
  进而导致整个工程编译失败（VBE 中该过程标红，`mod_date.Date_ToUtc` 报"方法或类成员未找到"），
  改用 `dt_local` / `utc_dt`

## [1.1.0] - 2026-09-26

### Added
- `mod_debug` 新增运行期日志：`Log_Debug/Info/Warn/Error`（时间戳+级别+过滤）、`Log_SetLevel`、
  `Log_GetLevel`、`Log_SetFile`（UTF-16 追加落盘，记事本可直接查看中文）；日志 IO 为"不吞错"原则的明确例外
- `JSON_Stringify` 支持美化输出：`pretty=True` 缩进换行（`indent` 控制空格），输出仍为合法 JSON 可再解析

### Fixed
- `JSON_Parse` 大整数（超出 Long 范围 2^31-1）不再抛 VBA 溢出错误：解析用 Decimal 承接，
  序列化对 Decimal 直出完整数字，大整数往返不丢精度
- `Folder_Ensure` 支持多级目录递归创建（原仅能创建单层）
- `mod_ui` 的 `File_Pick` / `Folder_Pick` 本地化 Office 常量并改晚绑定 `Object` 调用，
  兑现"无需勾选 Office 引用"承诺（模块版本 v1.1）

## [1.0.0] - 2026-09-26

### Added
- 纯层新增 7 个函数：`Array_IndexOf` / `Array_Remove` / `Array_RemoveAt`、`String_StartsWith` / `String_EndsWith` / `String_LeftPad`、`Dict_FromArray`（`Dict_To_Array` 之逆）
- 测试增至 106 例（纯层 99 + Excel 层 7）
- 新增 pre-push 静态门禁：契约检查（`vba_check_lint.py`）+ 工具脚本 `py_compile`，由 `tools/install_hooks.py` 安装

### Changed
- 新增 MIT License（`LICENSE`）
- `tools/verify_smoke.py` 去掉硬编码宿主路径，改为 `--host` 参数指定（缺省自动新建宿主并导入模块）

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