Attribute VB_Name = "mod_debug"
'=====================================================================
' mod_debug - 调试打印 / 运行期日志 通用工具（通用库 · v1.1）
'=====================================================================
' 目录 / Catalog
'   Show_Arr_Members(arr)   打印数组维度信息(类型/上下界/元素数)
'   Print_Array(arr)        打印数组元素(逗号分隔)
'   Print_Clc(col)          打印集合元素
'   Print_Dict(dict)        打印字典键值
'   Print_Lines(lines)      逐行打印字符串数组
'   Log_Debug(msg)          开发级日志（含时间戳/级别）
'   Log_Info(msg)           信息级日志
'   Log_Warn(msg)           警告级日志
'   Log_Error(msg)          错误级日志
'   Log_SetLevel(level)     设置最低输出级别："debug"/"info"/"warn"/"error"
'   Log_GetLevel() As String 当前最低级别名
'   Log_SetFile(path)       设置日志文件(UTF-16 追加，记事本可直接查看中文)；传空串关闭
' 说明：
'   - Print_* 仅输出到 VBE 立即窗口(Debug.Print)，用于开发期排查。
'   - Log_* 面向运行期：级别过滤后同时输出到立即窗口与日志文件（可选）。
'     级别由低到高 debug < info < warn < error，低于设定级别的不输出。
'   - 日志 IO 为尽力而为的辅助设施：文件写入失败静默（On Error 流控），
'     绝不影响主流程——这是本项目"不吞错"原则的明确例外。
'     写入采用 FSO 追加（UTF-16），每次调用开关句柄，高频循环场景慎用。
'=====================================================================
Option Explicit

Private Const LVL_DEBUG As Long = 0
Private Const LVL_INFO As Long = 1
Private Const LVL_WARN As Long = 2
Private Const LVL_ERROR As Long = 3

Private p_log_level As Long
Private p_log_file As String

'--- 开发期打印（VBE 立即窗口） ---------------------------------------

Public Sub Show_Arr_Members(ByRef arr As Variant)
    Dim nd As Long
    Dim lo As Long, hi As Long
    If Not IsArray(arr) Then
        Debug.Print "not an array, type=" & TypeName(arr)
        Exit Sub
    End If
    ' 逐维探测维度数与上下界：对不存在维度调用 LBound/UBound 会抛错，
    ' 用 On Error 作"到达最高维"的终止信号（这里是流控，非吞错）。
    On Error GoTo Done
    nd = 1
    Do While True
        lo = LBound(arr, nd)
        hi = UBound(arr, nd)
        Debug.Print "arr type=" & TypeName(arr); _
                "  dim[" & nd & "] " & lo & "~" & hi & " (" & (hi - lo + 1) & " elements)"
        nd = nd + 1
    Loop
Done:
    On Error GoTo 0
End Sub

Public Sub Print_Array(ByRef arr As Variant)
    Dim elem As Variant
    For Each elem In arr
        Debug.Print CStr(elem)
    Next elem
End Sub

Public Sub Print_Clc(ByRef col As Collection)
    Dim it As Variant
    For Each it In col
        Debug.Print CStr(it)
    Next it
End Sub

Public Sub Print_Dict(ByRef dict As Object)
    Dim k As Variant
    For Each k In dict.Keys
        Debug.Print CStr(k) & " => " & CStr(dict(k))
    Next k
End Sub

Public Sub Print_Lines(ByRef lines As Variant)
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        Debug.Print CStr(lines(i))
    Next i
End Sub

'--- 运行期日志 -------------------------------------------------------

Public Sub Log_Debug(ByVal msg As String)
    emit LVL_DEBUG, "DEBUG", msg
End Sub

Public Sub Log_Info(ByVal msg As String)
    emit LVL_INFO, "INFO", msg
End Sub

Public Sub Log_Warn(ByVal msg As String)
    emit LVL_WARN, "WARN", msg
End Sub

Public Sub Log_Error(ByVal msg As String)
    emit LVL_ERROR, "ERROR", msg
End Sub

' 设置最低输出级别。level 取值 "debug"/"info"/"warn"/"error"（大小写不敏感）。
Public Sub Log_SetLevel(ByVal level As String)
    Dim lv As Long
    lv = level_of(level)
    If lv < 0 Then Err.Raise 45000, "mod_debug", "Log_SetLevel: 未知级别 '" & level & "'"
    p_log_level = lv
End Sub

Public Function Log_GetLevel() As String
    Log_GetLevel = name_of(p_log_level)
End Function

' 设置日志文件（UTF-16 追加写）。传空串关闭（仅清路径，文件由 FSO 按次追加）。
Public Sub Log_SetFile(ByVal path As String)
    p_log_file = path
End Sub

'--- 内部实现 ---------------------------------------------------------

Private Sub emit(ByVal level As Long, ByVal tag As String, ByVal msg As String)
    Dim line As String
    If level < p_log_level Then Exit Sub
    line = "[" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "] [" & tag & "] " & msg
    Debug.Print line
    If Len(p_log_file) > 0 Then write_line line
End Sub

Private Function level_of(ByVal name As String) As Long
    Select Case LCase$(name)
        Case "debug": level_of = LVL_DEBUG
        Case "info":  level_of = LVL_INFO
        Case "warn":  level_of = LVL_WARN
        Case "error": level_of = LVL_ERROR
        Case Else:    level_of = -1
    End Select
End Function

Private Function name_of(ByVal level As Long) As String
    Select Case level
        Case LVL_DEBUG: name_of = "debug"
        Case LVL_INFO:  name_of = "info"
        Case LVL_WARN:  name_of = "warn"
        Case LVL_ERROR: name_of = "error"
        Case Else:      name_of = "?"
    End Select
End Function

Private Sub write_line(ByVal line As String)
    ' FSO 追加写（8=ForAppending；-1=TristateTrue 按 Unicode/UTF-16），
    ' 自动创建文件并在尾部续写，不会重复写 BOM。失败静默，尽力而为。
    On Error GoTo fail
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.OpenTextFile(p_log_file, 8, True, -1)
    ts.WriteLine line
    ts.Close
    Exit Sub
fail:
    On Error GoTo 0
End Sub
