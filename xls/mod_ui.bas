Attribute VB_Name = "mod_ui"
'=====================================================================
' mod_ui - Excel 交互对话框 通用工具（Excel 绑定层 · v1.1）
'=====================================================================
' 职责：封装文件/文件夹选择对话框。取消时返回空串。
'
' 目录 / Catalog
'   File_Pick([title, desc, pattern])  As String    对话框选文件
'   Folder_Pick([title])              As String    对话框选文件夹
' 说明：本模块使用 Application.FileDialog，强依赖 Excel 宿主，
'       归入 xls/ 层而非纯 VBA 的 src/。
'       对话框常量(msoFileDialog*)本地定义、句柄用 Object 晚绑定，
'       兑现"无需勾选外部引用"的承诺（不依赖 Office 类型库）。
'=====================================================================
Option Explicit

' msoFileDialogFilePicker=1 / msoFileDialogFolderPicker=2（Office 类型库取值）
Private Const MSO_FILE_DLG_PICKER As Long = 1
Private Const MSO_FILE_DLG_FOLDER As Long = 2

' 弹文件选择框（单选），返回所选文件全路径；取消返回空串。可用 filter_desc/pattern 限定类型。
Public Function File_Pick(Optional ByVal title As String = "选择文件", _
                          Optional ByVal filter_desc As String = "所有文件", _
                          Optional ByVal filter_pattern As String = "*.*") As String
    Dim dlg As Object
    Set dlg = Application.FileDialog(MSO_FILE_DLG_PICKER)
    With dlg
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filter_desc, filter_pattern
        If .Show = -1 Then File_Pick = .SelectedItems(1)
    End With
    Set dlg = Nothing
End Function

' 弹文件夹选择框，返回所选文件夹全路径；取消返回空串。
Public Function Folder_Pick(Optional ByVal title As String = "选择文件夹") As String
    Dim dlg As Object
    Set dlg = Application.FileDialog(MSO_FILE_DLG_FOLDER)
    With dlg
        .Title = title
        If .Show = -1 Then Folder_Pick = .SelectedItems(1)
    End With
    Set dlg = Nothing
End Function