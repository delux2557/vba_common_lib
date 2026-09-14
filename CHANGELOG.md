# Changelog

本项目遵循语义化版本（SemVer：`主.次.修订`）。机器可读的版本唯一来源为 `src/VERSION`，
每次发布前在此登记变更说明，并同步更新 `src/VERSION`。

## [Unreleased]

### 新增
- 新增 `mod_dict`：晚绑定 Dictionary 封装（Create/Set/Get/Exists/Count/Keys/Values/To_Array/Remove/Clear）。
- 新增 `mod_sort`：稳定归并排序 + 反转（升/降序、可选排序键回调，不改入参）。
- `mod_tests`：套件经 `run_protected` 容错分发，运行时错误不再中断整体，
  而是记录 `[CRASH] 套件名 ~ 最近步骤`（断言标签）后继续；断言计数 49 → 66 → 85。

### 修复
- `mod_sort`：对空数组 `Array()` 的 `ReDim(0 To -1)` 下标越界隐患（现返回空数组）。
- `mod_debug.Show_Arr_Members`：修正对一维数组的维度打印（原 `UBound(arr, 2)`
  对 1D 输入会抛错，导致 dims 信息错误）。
- 删除 `mod_tests` 中未被使用的死代码 `helper_gt_10`。
- 修正 `mod_file.File_Write` 误导注释：FSO `CreateTextFile` 第三参 `unicode=True`
  实际写入 UTF-16（非 UTF-8），与 `run_tests` 用 `utf-16` 解码保持一致。
- `mod_json`：修复对象引用传递。Dictionary 有带参默认属性 `.Item(Key)`，凡用 `=`（Let）
  承接包着对象的 Variant 会触发 `#450 错误的参数数`。解析器 `parse_value` 改为
  ByRef 输出并按 `Set`/`Let` 分流（对象引用用 `Set` 透传），`JSON_Parse` 依 `IsObject`
  决定用 `Set` 返回；`ser_dict` 形参 `ByRef→ByVal As Object` 修复序列化类型不符。
  `mod_tests` 的 suite_json 相应改为 `Set` 承接对象结果。

### 工具
- 新增 `tools/vba_check_lint.py`：离线静态契约检查（模块命名 / Catalog 双向一致 /
  Public 前缀归位 / 测试覆盖），不依赖 Excel，可作为 CI 快速门禁。
- `build_xlam` 构建时读取 `src/VERSION` 并注入到 xlam 自定义文档属性 `LibVersion`。

## [1.0.0] - 初始版本

### 新增
- 基础模块：`mod_array`（数组/集合·含高阶 map/filter/all/any）、`mod_string`、
  `mod_regex`（晚绑定正则）、`mod_date`、`mod_file`（FSO 晚绑定）、`mod_workbook`、
  `mod_range`、`mod_debug`、`mod_tests`（example-as-test 单测，46 条断言）。
- win32com 驱动真实 Excel 的构建/测试工具链：`run_tests.py`、`vba_import.py`、
  `vba_build_xlam.py`、`vba_new_host.py`、`vba_excel.py`。
- 命名双轨制：对外 API 用 `PascalCase + 域名前缀`，内部实现用 `snake_case`。