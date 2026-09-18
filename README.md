# VBA 通用工具库（VBA Common Lib）

把散落在各个工作簿里、可复用的 VBA 函数提取成一套**源码即版本管理**的通用库。
支持两种使用形态：

1. **导入式**：一键把 `src/*.bas` 同步进任意 `.xlsm`（开发/轻量用）。
2. **加载项式**：打包成 `.xlam`，多个工作簿通过 VBE 引用一次，升级全端生效（推荐交付）。

---

## 目录结构

```
vba_common_lib/
├── src/                  # 纯 VBA 层（宿主无关，进 git）；模块名 mod_*
│   ├── mod_array.bas
│   ├── mod_string.bas
│   ├── mod_regex.bas
│   ├── mod_date.bas      # Date_Stamp(kind) 统一时间戳
│   ├── mod_file.bas      # FSO 文件/文件夹（无 UI）
│   ├── mod_debug.bas
│   ├── mod_dict.bas      # P1：晚绑定 Dictionary 封装
│   ├── mod_sort.bas      # P1：稳定排序 / 反转
│   ├── mod_json.bas      # P1：JSON 解析 / 序列化
│   └── mod_tests.bas     # 纯层单元测试（TDD：新函数在此补断言）
├── xls/                  # Excel 绑定层（依赖 Range/Workbook/Application.FileDialog）
│   ├── mod_workbook.bas
│   ├── mod_range.bas
│   ├── mod_ui.bas        # 文件对话框 File_Pick / Folder_Pick
│   └── mod_tests_excel.bas  # Excel 层单元测试
├── src/VERSION           # 语义化版本号唯一来源（如 1.0.0）；构建时注入 mod_version
├── CHANGELOG.md          # 版本变更记录（Keep a Changelog 风格）
├── tools/                # Python（win32com 驱动真实 Excel）
│   ├── vba_excel.py      # 共享封装（启动/清空/导入/打包，支持多 --src 分层）
│   ├── vba_import.py     # 一键 清空+导入 到目标 .xlsm
│   ├── run_tests.py      # 跑单测（--entry 选纯层/Excel 层），返回失败数
│   ├── vba_build_xlam.py # 打包 .xlam
│   ├── vba_check_lint.py # 离线静态契约检查（可 CI 门禁）
│   └── vba_new_host.py   # 新建空宿主 .xlsm（备用）
├── build/                # 产物（测试宿主 / .xlam）
└── README.md
```

> **分层原则**：`src/` 只含不依赖 Excel 对象模型的纯 VBA，任何宿主都能编译运行，
> 可单独打 xlam / 独立测试；`xls/` 绑定 `Range` / `Workbook` / `Application`，
> 按需随纯层一起导入。`xls/` 可依赖 `src/`，反之不可。

---

## 功能清单（60 个对外 API · 12 个功能模块 = 纯层 9 + Excel 层 3）

### 纯 VBA 层 `src/`（宿主无关）

### 数组运算 `mod_array`
`Array_Contains` / `Array_Append` / `Array_Extend` / `Array_Distinct`（保持首见顺序）
/ `Array_Map` / `Array_Filter` / `Array_All` / `Array_Any`（高阶回调，走 `Application.Run`）
/ `Collection_To_Array` / `Array_To_Collection`

### 字符串 `mod_string`
`String_To_Array` / `String_Join` / `String_Trim`（可指定字符集）/ `String_LTrim` / `String_RTrim` / `String_Random`（长度+字符集可选）

### 正则 `mod_regex`（晚绑定）
`Regex_Test` / `Regex_Find` / `Regex_Replace`（全局）

### 日期时间 `mod_date`
`Date_Stamp(kind)`：`date`(YYYYMMDD) / `time`(hhmmss) / `datetime`(YYYY_MMDD_hhmm) / `stamp`(YYYY_MMDD_hhmmss, 默认，文件名安全)

### 文件路径 `mod_file`（FSO 晚绑定）
存在判断：`File_Exists` / `Folder_Exists`；文件操作：`Folder_Ensure` / `File_Copy` / `File_Write`（注意默认 `unicode=True` 为 UTF-16）
路径解析：`File_Name` / `File_BaseName` / `File_ExtName` / `File_Name_Valid`
`Folder_ListFiles`（可递归）

### 调试输出 `mod_debug`
`Show_Arr_Members`（维度）/ `Print_Array` / `Print_Clc` / `Print_Dict` / `Print_Lines`

### 字典封装 `mod_dict`（晚绑定）
`Dict_Create` / `Dict_Set` / `Dict_Get`（可带默认值）/ `Dict_Exists` / `Dict_Count`
/ `Dict_Keys` / `Dict_Values` / `Dict_To_Array`（二维）/ `Dict_Remove` / `Dict_Clear`

### 排序 `mod_sort`
`Array_Sort`（稳定归并，可选键回调，返回新数组）/ `Array_Reverse`

### JSON `mod_json`（纯手写递归下降）
`JSON_Parse`（对象→`Dictionary`，数组→0 基数组）/ `JSON_Stringify`（序列化）
> **注意**：`JSON_Parse` 返回对象时须用 `Set` 承接（`Set o = mod_json.JSON_Parse(...)`）；
> 字典有带参默认属性 `.Item`，用 `=` 会触发 `#450`。

### 版本 / 测试
`mod_version`（由 `src/VERSION` 自动生成）· `mod_tests`（TDD 断言框架，纯层套件经 `run_protected` 容错分发）。

### Excel 绑定层 `xls/`（依赖 Excel 宿主，随纯层一起导入）

### 工作簿 `mod_workbook`
`Sheet_Exists` / `Workbook_Exists` / `Workbook_SheetNames` / `Workbook_AllSheetNames` / `Workbook_SaveWithBackup`（带时间戳备份；`do_save` 可选，出错时自动恢复宿主 `DisplayAlerts`）

### 单元格区域 `mod_range`
`Range_Address` / `Range_SheetAddress`（带表名）/ `Range_ShiftAddress`（偏移）/ `Range_JoinedAddress`（多区域拼接）

### 对话框 `mod_ui`
`File_Pick` / `Folder_Pick`（`Application.FileDialog`）

---

## 命名规范（双轨制）

- **对外 API（Public，会被 Excel 单元格 `mod_*.xxx` 直接调用 / 做成 UDF）**
  ：`PascalCase + 域名前缀`。例：`Array_Contains`、`Sheet_Exists`。
- **内部实现（Private 过程、局部变量、模块级变量）：`snake_case`**（Python 手感）。
  例：`src_arr`、`seen_keys`、`m_config`。
- **模块名：`mod_<域>`**，作为调用限定符 `mod_array.Array_Contains`。
- **例外（不可改）**：`Workbook_Open` / `Sheet_Select` 等事件名、VBA 内建名。
- **红线**：UDF 只能由「标准模块里的 Public 函数」承担；类模块方法、带参数引用无法进单元格。

## 函数名体系（新函数按前缀归位）

VBA 没有命名空间，「体系」靠 **域名前缀字典 + 模块顶部目录注释** 落地：

| 前缀 | 含义 | 后缀约定 |
|---|---|---|
| `Array_*` | 数组操作（`Array_Contains/Distinct/Extend/Append`） | |
| `Array_Map/Filter/All/Any` | 高阶函数（回调传 `Application.Run` 可达的函数名） | |
| `Collection_*` / `*_To_*` | 集合操作 / 转换（`Collection_To_Array`、`String_To_Array`） | 转换统一 `<From>_To_<To>` |
| `String_*` | 字符串操作（`String_Trim/Join/Random`） | |
| `Regex_*` | 正则（`Regex_Test/Find/Replace`） | |
| `File_*` / `Folder_*` | 文件 / 文件夹（`File_Exists`、`Folder_Ensure`、`Folder_ListFiles`） | |
| `Date_*` | 日期时间（`Date_Stamp`） | |
| `Sheet_*` / `Workbook_*` | 工作表 / 工作簿（`Sheet_Exists`、`Workbook_SheetNames`） | |
| `Range_*` | Range 地址（`Range_Address`、`Range_ShiftAddress`） | |

**谓词/辅助后缀**：`_Exists`（`Folder_Exists`）、`_Contains`（`Array_Contains`）、`_IsXxx`（`File_Name_Valid`）。
**原材料规范**：谓词返回 `Boolean`；转换返回 `Variant` 数组；对象转换返回 `Collection`。

**新增函数 Checklist（体系强制项）**：
1. 按域名前缀归位到对应 `mod_*`；2. `Public` 进标准模块；3. 在 `mod_tests` 补一条断言；4. 登记进该模块顶部目录注释。

---

## 使用流程

### 0. 前置：开启 COM 自动化
自动脚本需要访问 VB 工程，请开启：
**文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 勾选『信任对 VBA 工程对象模型的访问』**，
并重启 Excel。另需 `pip install pywin32`。

### 1. 跑单元测试（推荐每改一次就测一遍）
```bash
# 纯 VBA 层（默认 --entry mod_tests）
python tools/run_tests.py --src src --build build
# Excel 绑定层（需同时导入 src+xls，入口 mod_tests_excel）
python tools/run_tests.py --src src --src xls --entry mod_tests_excel --build build
```
退出码 = 失败数；0 表示全部通过。

### 1b. 离线静态契约检查（不依赖 Excel，可进 CI）
```bash
python tools/vba_check_lint.py --src src --src xls
```

### 2. 一键同步到某个工作簿
```bash
# 纯层
python tools/vba_import.py --src src --target build/VBA_Common_test.xlsm
# 纯层 + Excel 绑定层
python tools/vba_import.py --src src --src xls --target build/VBA_Common_test.xlsm
# 想保留某些模块不删：--keep mod_importer
```

### 3. 打包成加载项
```bash
python tools/vba_build_xlam.py --src src --out build/VBA_Common.xlam
python tools/vba_build_xlam.py --src src --src xls --out build/VBA_Common_Full.xlam
```
宿主工作簿：VBE → 工具 → 引用 → 勾选该加载项 → 直接 `mod_array.Array_Contains(...)`。

---

## 设计要点

- **不吞错**：清除了源码中散布的 `On Error Resume Next`，错误尽早暴露；
  空数组/空集合一律返回 `Array()`，谓词对空输入返回 `False`。
- **晚绑定**：正则、字典、FSO 均 `CreateObject` 晚绑定，**无需手动勾选外部引用**。
- **无共享状态**：正则每次新建 RegExp；FSO 用模块级单例惰性初始化。
- **测试即文档**：`Run_All_Tests` 的每条断言即每个函数的可运行用法（example-as-test）。