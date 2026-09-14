# -*- coding: utf-8 -*-
"""VBA Common Lib - Excel COM 共享封装

包含：启动/关闭 Excel、访问 VBA 工程、清空标准模块、批量导入 .bas、
新建宿主工作簿、打包 .xlam 等公共能力。所有改动都用外部 COM 驱动真实 Excel，
这才是“源码目录 + 一键同步入工作簿/加载项”的标准做法。
"""
from __future__ import annotations

import io
import os
import sys
import tempfile

try:
    import win32com.client as win32
    from win32com.client import constants
except ImportError:
    sys.exit("缺少 pywin32，请先执行: pip install pywin32")

# VBE 组件类型
TYP_DOCUMENT = 100          # ThisWorkbook / Sheet 代码页
TYP_STD_MODULE = 1
TYP_CLASS_MODULE = 2
TYP_FORM = 3

# Excel 文件格式
FORMAT_XLSM = 52            # xlOpenXMLWorkbookMacroEnabled
FORMAT_XLAM = 55            # xlOpenXMLAddIn

# 进程内单例 Excel（模块级共享，避免反复启停实例）
_excel_ref: dict = {"app": None, "excel": None}


def _proc_pid() -> str:
    try:
        import _winapi
        return str(_winapi.GetCurrentProcessId())
    except Exception:
        return "com"


def start_excel(visible: bool = False):
    """启动（或复用）Excel 实例。用本地 PID 标记进程内 Excel，避免跨调用重复创建。"""
    if _excel_ref["app"] is not None:
        return _excel_ref["app"]
    if "excel" not in dir(win32):
        pass  # module 原生存在
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False
    app = excel.Application
    _excel_ref["app"] = app
    _excel_ref["excel"] = excel
    return app


def quit_excel():
    app = _excel_ref.get("excel")
    if app is None:
        return
    try:
        app.DisplayAlerts = False
        if app.Workbooks.Count:
            app.Workbooks.Close()
        app.Quit()
    except Exception:
        pass
    _excel_ref["app"] = None
    _excel_ref["excel"] = None


def trust_ok(excel) -> bool:
    """是否已勾选『信任对 VBA 工程对象模型的访问』。"""
    try:
        _ = excel.ActiveWorkbook.VBProject.Name
        return True
    except Exception:
        return False


def _vb_project(wb):
    """获取工作簿的 VBProject；未授权 Trust 时抛出带指引的异常。"""
    try:
        return wb.VBProject
    except Exception:
        sys.exit(
            "\n[错误] 无法访问 VBProject。请先在 Excel 中开启：\n"
            "   文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置\n"
            "   ⊙ 勾选『信任对 VBA 工程对象模型的访问』并重启 Excel。\n"
            "（该设置由 VBA 工程对象模型访问接口决定，任何自动化脚本都需要它。）"
        )


def list_components(wb):
    return [c for c in _vb_project(wb).VBComponents]


def clear_modules(wb, keep_names=()):
    """删除工程中所有『非文档』组件(标准模块/类模块/窗体)，保留 ThisWorkbook/Sheet 及 keep_names。"""
    keep = set(keep_names or [])
    comps = list_components(wb)
    removed = []
    for comp in comps:
        # TYP_DOCUMENT=100：ThisWorkbook / Worksheet 代码页，永远保留
        if comp.Type == TYP_DOCUMENT:
            continue
        name = comp.Name
        if name in keep:
            continue
        try:
            # 注意：删除后 comp 句柄即失效，故先取 name 再 Remove
            wb.VBProject.VBComponents.Remove(comp)
            removed.append(name)
        except Exception as exc:  # noqa: BLE001
            print(f"  [warn] 删除组件 {name} 失败: {exc}")
    return removed


def _system_ansi_codec():
    """返回 Excel 期望的源码编码：本机 ANSI 代码页。
    repo 里源码存 UTF-8，这里在导入前转成系统 ANSI，避免中文乱码/引号断裂导致编译错误。
    """
    if os.name != "nt":
        return "utf-8"
    try:
        import ctypes
        acp = ctypes.windll.kernel32.GetACP()
    except Exception:  # noqa: BLE001
        acp = 65001
    # 65001 = UTF-8，无需转换；其余按 cp<代码页> 处理（本机一般为 936 = GBK）
    return "utf-8" if acp == 65001 else f"cp{acp:02d}"


def _write_import_tmp(name, text, codec):
    """规范换行并按目标编码写临时 .bas，返回路径；交给 Excel 正确解析中文。"""
    text = text.replace("\r\n", "\n").replace("\n", "\r\n")
    tmp = os.path.join(tempfile.mkdtemp(prefix="vbaimp_"), name)
    enc = "utf-8" if (codec == "utf-8") else codec
    with io.open(tmp, "w", encoding=enc, newline="") as fh:
        fh.write(text)
    return tmp


# 版本查询模块模板：构建/导入时由 src/VERSION 生成，【不在 src 落盘】，
# 保证版本唯一来源是 src/VERSION，产物可重建且运行期可用 mod_version.Version_String() 查询。
VERSION_MODULE_NAME = "mod_version"
VERSION_MODULE_TEMPLATE = (
    'Attribute VB_Name = "mod_version"\n'
    "Attribute VB_Exposed = True\n"
    "'=====================================================================\n"
    "' mod_version - 版本查询（构建期由 src/VERSION 自动生成，请勿手改）\n"
    "'=====================================================================\n"
    "Option Explicit\n"
    "\n"
    "Public Const LIB_VERSION As String = \"__VERSION__\"\n"
    "\n"
    "Public Function Version_String() As String\n"
    "    Version_String = LIB_VERSION\n"
    "End Function\n"
)


def version_module(version):
    """用版本号填充模板，返回 mod_version.bas 的正文。"""
    return VERSION_MODULE_TEMPLATE.replace("__VERSION__", version or "0.0.0")


def import_bas(wb, src_dir, pattern=".bas"):
    """按文件名排序，逐个 Import 指定目录下的模块文件。返回导入列表。

    源码为 UTF-8，导入前自动转写成系统 ANSI 临时文件，交给 Excel 正确解析中文。
    此外会用 src/VERSION 追加生成一个 mod_version 版本查询模块（单一来源，不在 src 落盘）。
    """
    names = []
    if not os.path.isdir(src_dir):
        sys.exit(f"[错误] 源码目录不存在: {src_dir}")
    codec = _system_ansi_codec()
    files = sorted(f for f in os.listdir(src_dir) if f.lower().endswith(pattern))
    for fn in files:
        if codec == "utf-8":
            wb.VBProject.VBComponents.Import(os.path.join(src_dir, fn))
        else:
            with io.open(os.path.join(src_dir, fn), "r", encoding="utf-8-sig") as fh:
                text = fh.read()
            tmp = _write_import_tmp(fn, text, codec)
            wb.VBProject.VBComponents.Import(tmp)
        names.append(fn)
    # 追加版本查询模块
    try:
        vt = version_module(read_version(src_dir))
        tmp = _write_import_tmp(VERSION_MODULE_NAME + ".bas", vt, codec)
        wb.VBProject.VBComponents.Import(tmp)
        names.append(VERSION_MODULE_NAME + ".bas")
    except Exception:  # noqa: BLE001  版本模块缺失时不影响主流程
        pass
    return names


def new_host(host_path, sheet_name="Sheet1"):
    """新建一个空 .xlsm 宿主（含一个命名工作表），返回工作簿对象。"""
    app = start_excel()
    wb = app.Workbooks.Add()
    ws = wb.Worksheets(1)
    ws.Name = sheet_name
    os.makedirs(os.path.dirname(os.path.abspath(host_path)), exist_ok=True)
    wb.SaveAs(host_path, FORMAT_XLSM)
    return wb


def open_or_create_host(host_path):
    app = start_excel()
    if os.path.exists(host_path):
        return app.Workbooks.Open(host_path)
    return new_host(host_path)


def read_version(src_dir, default="0.0.0"):
    """读取 src/VERSION 作为唯一版本来源；缺失时返回 default。"""
    vfile = os.path.join(src_dir, "VERSION")
    if not os.path.isfile(vfile):
        return default
    with io.open(vfile, "r", encoding="utf-8") as fh:
        return fh.read().strip() or default


def build_xlam(src_dir, out_path, display_name="VBA_Common"):
    """把 src 下的 .bas 打进一个 .xlam 加载项。"""
    app = start_excel()
    wb = app.Workbooks.Add()
    # 工作簿必须保留至少一张工作表；裁到只剩 1 张并隐藏，供自检宿主使用
    while wb.Worksheets.Count > 1:
        wb.Worksheets(wb.Worksheets.Count).Delete()
    wb.Worksheets(1).Visible = -1  # xlSheetHidden
    # 说明：ThisWorkbook / 工作表等『文档』代码页不能也不需删除，保持为空即可
    import_bas(wb, src_dir)  # 含由 src/VERSION 自动生成的 mod_version
    # 加载项属性
    try:
        wb.IsAddIn = True
    except Exception:
        pass
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    wb.SaveAs(out_path, FORMAT_XLAM)
    wb.Close(SaveChanges=False)
    return out_path