Attribute VB_Name = "mod_ui"
'=====================================================================
' mod_ui - Excel 交互对话框 通用工具（Excel 绑定层 · v1.0）
'=====================================================================
' 目录 / Catalog
'   File_Pick([title, desc, pattern])  As String    对话框选文件
'   Folder_Pick([title])              As String    对话框选文件夹
' 说明：本模块使用 Application.FileDialog，强依赖 Excel 宿主，
'       归入 xls/ 层而非纯 VBA 的 src/。
'=====================================================================
Option Explicit

Public Function File_Pick(Optional ByVal title As String = "选择文件", _
                          Optional ByVal filter_desc As String = "所有文件", _
                          Optional ByVal filter_pattern As String = "*.*") As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add filter_desc, filter_pattern
        If .Show = -1 Then File_Pick = .SelectedItems(1)
    End With
    Set dlg = Nothing
End Function

Public Function Folder_Pick(Optional ByVal title As String = "选择文件夹") As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = title
        If .Show = -1 Then Folder_Pick = .SelectedItems(1)
    End With
    Set dlg = Nothing
End Function