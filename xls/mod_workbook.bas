Attribute VB_Name = "mod_workbook"
'=====================================================================
' mod_workbook - 工作簿 / 工作表 通用工具（通用库 · v1.0）
'=====================================================================
' 职责：Excel 工作簿/工作表查询与安全备份。依赖 Excel 宿主，归入 xls/ 绑定层。
' 约定：wb Is Nothing 一律安全返回（False / Array()），不抛错。
'
' 目录 / Catalog
'   Sheet_Exists(sheet_name, wb)          As Boolean   工作表是否存在于指定工作簿
'   Sheet_Get(sheet_name, wb)             As Worksheet 按名取工作表对象(不存在返回 Nothing)
'   Sheet_Create(name, wb[, before])      As Worksheet 新建工作表(重名抛错 #45000)
'   Sheet_Ensure(name, wb[, before])      As Worksheet 幂等获取:不存在则创建,存在则返回已有
'   Sheet_Rename(sh, new_name)                       重命名(空名/重名抛错 #45000)
'   Sheet_Copy(sh[, new_name])            As Worksheet 复制工作表(默认 "原名 (2)",恢复活动表)
'   Sheet_Delete(sh)                                 删除工作表(至少保留一张,临时关 DisplayAlerts 并恢复)
'   Workbook_SheetCount(wb)               As Long      工作表数量(wb 为空返回 0)
'   Workbook_Exists(wbname[, fmt])        As Boolean   工作簿是否已打开(按名,忽略大小写)
'   Workbook_SheetNames(wb)               As Variant   列出指定工作簿全部工作表名(数组)
'   Workbook_AllSheetNames()              As Variant   列出所有打开工作簿的工作表名(数组)
'   Workbook_SaveWithBackup(wb, folder[, prefix, do_save])  保存并备份副本(带时间后缀,确保恢复 DisplayAlerts)
'=====================================================================
Option Explicit

' 指定工作簿里是否存在名为 sheet_name 的工作表。
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

' 按名称判断工作簿是否已打开；可传带/不带扩展名（fmt 默认 xlsx），忽略大小写。
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

' 列出指定工作簿的全部工作表名（0 基数组）；wb 为空返回 Array()。
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

' 列出所有打开工作簿的全部工作表名，格式 "工作簿名!工作表名"。
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
                                   Optional ByVal prefix As String = "backup_", _
                                   Optional ByVal do_save As Boolean = True)
    Dim ext As String
    Dim dest As String
    Dim prev_display As Boolean
    If wb Is Nothing Then Exit Sub
    ext = mod_file.File_ExtName(wb.Name)
    dest = backup_folder & "\" & prefix & mod_date.Date_Stamp() & "." & ext
    mod_file.Folder_Ensure backup_folder
    ' 暂缓警告并保证无论成败都恢复宿主 DisplayAlerts，避免污染调用方状态
    prev_display = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo fail
    If do_save Then wb.Save
    wb.SaveCopyAs dest
    Application.DisplayAlerts = prev_display
    Exit Sub
fail:
    Application.DisplayAlerts = prev_display
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function StrCaseEq(ByVal a As String, ByVal b As String) As Boolean
    StrCaseEq = (StrComp(a, b, vbTextCompare) = 0)
End Function

'--- 工作表批量管理（建 / 取 / 改名 / 复制 / 删除） ---------------------

' 按名称取工作表对象，省去 wb.Sheets(name) 前手动判存在的写法；
' 表不存在时返回 Nothing（调用方自行判断），wb 为空同样返回 Nothing。
Public Function Sheet_Get(ByVal sheet_name As String, ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Exit Function
    If Sheet_Exists(sheet_name, wb) Then Set Sheet_Get = wb.Sheets(sheet_name)
End Function

' 新建工作表并返回它。默认插到末尾；传 before（表名或 Worksheet 对象）则插到其前。
' 表名已存在或为空时抛错 #45000，避免静默覆盖已有数据。
Public Function Sheet_Create(ByVal name As String, ByVal wb As Workbook, _
                             Optional ByVal before_sheet As Variant) As Worksheet
    Dim new_ws As Worksheet
    If wb Is Nothing Then Err.Raise 45000, "mod_workbook", "Sheet_Create: wb 为 Nothing"
    If Len(name) = 0 Then Err.Raise 45000, "mod_workbook", "Sheet_Create: 表名不能为空"
    If Sheet_Exists(name, wb) Then _
        Err.Raise 45000, "mod_workbook", "Sheet_Create: 已存在同名表 '" & name & "'"
    If IsMissing(before_sheet) Then
        Set new_ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    Else
        Set new_ws = wb.Sheets.Add(Before:=resolve_sheet(wb, before_sheet))
    End If
    new_ws.Name = name
    Set Sheet_Create = new_ws
End Function

' 幂等获取工作表：存在则返回已有表，不存在则新建——"确保某张表在"的日常写法。
' 参数与 Sheet_Create 相同；创建行为也一致（重名由内部 Sheet_Create 兜底，不会真重名）。
Public Function Sheet_Ensure(ByVal name As String, ByVal wb As Workbook, _
                             Optional ByVal before_sheet As Variant) As Worksheet
    If Sheet_Exists(name, wb) Then
        Set Sheet_Ensure = wb.Sheets(name)
    Else
        Set Sheet_Ensure = Sheet_Create(name, wb, before_sheet)
    End If
End Function

' 重命名工作表。空名 / 与其他表重名时抛错 #45000；
' 新名与当前名相同则直接返回（幂等，避免无谓操作）。
Public Sub Sheet_Rename(ByVal sh As Worksheet, ByVal new_name As String)
    Dim wb As Workbook
    If sh Is Nothing Then Err.Raise 45000, "mod_workbook", "Sheet_Rename: sh 为 Nothing"
    If Len(new_name) = 0 Then Err.Raise 45000, "mod_workbook", "Sheet_Rename: 新表名不能为空"
    If sh.Name = new_name Then Exit Sub
    Set wb = sh.Parent
    If Sheet_Exists(new_name, wb) Then _
        Err.Raise 45000, "mod_workbook", "Sheet_Rename: 已存在同名表 '" & new_name & "'"
    sh.Name = new_name
End Sub

' 复制工作表并返回副本。不传 new_name 时用 Excel 默认名 "原名 (2)"；
' 复制完成后恢复原活动表（Excel 默认会跳到副本，库函数保持"无副作用"）。
Public Function Sheet_Copy(ByVal sh As Worksheet, Optional ByVal new_name As String) As Worksheet
    Dim wb As Workbook
    Dim saved_active As Object
    Dim dup As Worksheet
    If sh Is Nothing Then Err.Raise 45000, "mod_workbook", "Sheet_Copy: sh 为 Nothing"
    Set wb = sh.Parent
    If Len(new_name) > 0 And Sheet_Exists(new_name, wb) Then _
        Err.Raise 45000, "mod_workbook", "Sheet_Copy: 已存在同名表 '" & new_name & "'"
    Set saved_active = wb.ActiveSheet
    sh.Copy After:=sh
    Set dup = wb.ActiveSheet
    If Len(new_name) > 0 Then dup.Name = new_name
    saved_active.Activate
    Set Sheet_Copy = dup
End Function

' 删除工作表。破坏性操作，调用方确认后再调；
' 工作簿只剩一张表时抛错 #45000（Excel 本身禁止删掉最后一张表）。
' 删除期间临时关 DisplayAlerts 免去确认弹窗，无论成败都恢复原设置。
Public Sub Sheet_Delete(ByVal sh As Worksheet)
    Dim wb As Workbook
    Dim prev_display As Boolean
    If sh Is Nothing Then Err.Raise 45000, "mod_workbook", "Sheet_Delete: sh 为 Nothing"
    Set wb = sh.Parent
    If wb.Sheets.Count <= 1 Then _
        Err.Raise 45000, "mod_workbook", "Sheet_Delete: 工作簿至少需保留一张表"
    prev_display = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo fail
    sh.Delete
    Application.DisplayAlerts = prev_display
    Exit Sub
fail:
    Application.DisplayAlerts = prev_display
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' 统计工作簿里工作表数量；wb 为空返回 0（与 Sheet_Exists 的"安全返回"约定一致）。
Public Function Workbook_SheetCount(ByVal wb As Workbook) As Long
    If wb Is Nothing Then Exit Function
    Workbook_SheetCount = wb.Sheets.Count
End Function

' 私有辅助：把"表名(字符串)或 Worksheet 对象"统一解析成 Worksheet；
' 供 before 类参数复用。无效输入抛错 #45000。
Private Function resolve_sheet(ByVal wb As Workbook, ByVal spec As Variant) As Worksheet
    If TypeName(spec) = "Worksheet" Then
        Set resolve_sheet = spec
    ElseIf VarType(spec) = vbString Then
        If Sheet_Exists(spec, wb) Then
            Set resolve_sheet = wb.Sheets(spec)
        Else
            Err.Raise 45000, "mod_workbook", "resolve_sheet: 不存在工作表 '" & spec & "'"
        End If
    Else
        Err.Raise 45000, "mod_workbook", "resolve_sheet: 需传表名(字符串)或 Worksheet 对象"
    End If
End Function