Attribute VB_Name = "mod_workbook"
'=====================================================================
' mod_workbook - 工作簿 / 工作表 通用工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   Sheet_Exists(sheet_name, wb)          As Boolean   工作表是否存在于指定工作簿
'   Workbook_Exists(wbname[, fmt])        As Boolean   工作簿是否已打开(按名,忽略大小写)
'   Workbook_SheetNames(wb)               As Variant   列出指定工作簿全部工作表名(数组)
'   Workbook_AllSheetNames()              As Variant   列出所有打开工作簿的工作表名(数组)
'   Workbook_SaveWithBackup(wb, folder[, prefix])  保存并备份副本(带时间后缀)
'=====================================================================
Option Explicit

Public Function Sheet_Exists(ByVal sheet_name As String, ByRef wb As Workbook) As Boolean
    Dim sh As Variant
    Sheet_Exists = False
    If wb Is Nothing Then Exit Function
    For Each sh In wb.Sheets
        If sh.Name = sheet_name Then
            Sheet_Exists = True
            Exit Function
        End If
    Next sh
End Function

Public Function Workbook_Exists(ByVal wb_name As String, Optional ByVal fmt As String = "xlsx") As Boolean
    Dim wb As Workbook
    Workbook_Exists = False
    For Each wb In Application.Workbooks
        If StrCaseEq(wb.Name, wb_name) Or StrCaseEq(wb.Name, wb_name & "." & fmt) Then
            Workbook_Exists = True
            Exit For
        End If
    Next wb
End Function

Public Function Workbook_SheetNames(ByRef wb As Workbook) As Variant
    Dim col As New Collection
    Dim sh As Variant
    If wb Is Nothing Then
        Workbook_SheetNames = Array()
        Exit Function
    End If
    For Each sh In wb.Sheets
        col.Add sh.Name
    Next sh
    Workbook_SheetNames = mod_array.Collection_To_Array(col)
End Function

Public Function Workbook_AllSheetNames() As Variant
    Dim wb As Workbook
    Dim col As New Collection
    For Each wb In Application.Workbooks
        Dim sh As Variant
        For Each sh In wb.Sheets
            col.Add wb.Name & "!" & sh.Name
        Next sh
    Next wb
    Workbook_AllSheetNames = mod_array.Collection_To_Array(col)
End Function

' 先保存工作簿，再将副本备份到 backup_folder（文件名带时间戳，前缀可配）
Public Sub Workbook_SaveWithBackup(ByRef wb As Workbook, ByVal backup_folder As String, _
                                   Optional ByVal prefix As String = "backup_")
    Dim ext As String
    Dim dest As String
    If wb Is Nothing Then Exit Sub
    mod_file.Folder_Ensure backup_folder
    ext = mod_file.File_ExtName(wb.Name)
    dest = backup_folder & "\" & prefix & mod_date.Full_DateTime_String & "." & ext
    Application.DisplayAlerts = False
    wb.Save
    wb.SaveCopyAs dest
    Application.DisplayAlerts = True
End Sub

Private Function StrCaseEq(ByVal a As String, ByVal b As String) As Boolean
    StrCaseEq = (StrComp(a, b, vbTextCompare) = 0)
End Function