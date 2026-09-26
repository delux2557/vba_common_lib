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
'   Repr(obj[, max_depth])  通用对象字符串表示(Python 风格 repr：数组/集合/字典/Range/普通对象)
'   Print_Repr(obj[, ...])  Debug.Print Repr 的便捷入口
'   Log_Debug(msg)          开发级日志（含时间戳/级别）
'   Log_Info(msg)           信息级日志
'   Log_Warn(msg)           警告级日志
'   Log_Error(msg)          错误级日志
'   Log_SetLevel(level)     设置最低输出级别："debug"/"info"/"warn"/"error"
'   Log_GetLevel() As String 当前最低级别名
'   Log_SetFile(path)       设置日志文件(UTF-16 追加，记事本可直接查看中文)；传空串关闭
' 说明：
'   - Print_* 仅输出到 VBE 立即窗口(Debug.Print)，用于开发期排查。
'   - Repr 返回任意值的字符串表示：Debug.Print 直接打印对象会抛错(#438 等)，
'     Repr 按类型分发给出 Python 风格展示；标量字符串自动加引号并转义内部引号。
'     深度默认 16 层，循环引用在深度处截断为 "..."，不会无限递归。
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

'--- 通用对象表示（Python 风格 repr） -----------------------------------

Public Function Repr(obj As Variant, Optional ByVal max_depth As Long = 16) As String
    If max_depth <= 0 Then
        Repr = "..."
        Exit Function
    End If
    If IsObject(obj) Then
        If obj Is Nothing Then
            Repr = "<Nothing>"
        Else
            Repr = repr_object(obj, max_depth)
        End If
    ElseIf IsArray(obj) Then
        Repr = repr_array(obj, max_depth)
    ElseIf IsEmpty(obj) Then
        Repr = "Empty"
    ElseIf IsNull(obj) Then
        Repr = "Null"
    ElseIf IsError(obj) Then
        Repr = "Error(" & CStr(CLng(obj)) & ")"
    ElseIf VarType(obj) = vbString Then
        Repr = """" & Replace$(CStr(obj), """", """""") & """"
    Else
        Repr = CStr(obj)
    End If
End Function

Public Sub Print_Repr(obj As Variant, Optional ByVal max_depth As Long = 16)
    Debug.Print Repr(obj, max_depth)
End Sub

Private Function repr_object(ByVal obj As Object, ByVal depth As Long) As String
    Dim tn As String
    Dim s As String
    tn = TypeName(obj)
    Select Case tn
        Case "Collection"
            repr_object = "Collection(" & repr_collection(obj, depth) & ")"
        Case "Dictionary"
            repr_object = "Dictionary(" & repr_dictionary(obj, depth) & ")"
        Case "Range"
            repr_object = "Range(""" & obj.Address(False, False) & """, Value=" & _
                          Repr(obj.Value, depth - 1) & ")"
        Case Else
            ' 有默认成员的对象按默认值展示（近似 Python 的 str()）；否则退化 <类型 at 地址>
            On Error Resume Next
            s = CStr(obj)
            If Err.Number = 0 Then
                On Error GoTo 0
                repr_object = s
            Else
                On Error GoTo 0
                repr_object = "<" & tn & " at 0x" & Hex$(ObjPtr(obj)) & ">"
            End If
    End Select
End Function

Private Function repr_collection(ByVal col As Collection, ByVal depth As Long) As String
    Dim it As Variant
    Dim parts() As String
    Dim n As Long
    If col.Count = 0 Then
        repr_collection = ""
        Exit Function
    End If
    ReDim parts(0 To col.Count - 1)
    n = 0
    For Each it In col
        parts(n) = Repr(it, depth - 1)
        n = n + 1
    Next it
    repr_collection = Join(parts, ", ")
End Function

Private Function repr_dictionary(ByVal d As Object, ByVal depth As Long) As String
    Dim k As Variant
    Dim parts() As String
    Dim n As Long
    If d.Count = 0 Then
        repr_dictionary = ""
        Exit Function
    End If
    ReDim parts(0 To d.Count - 1)
    n = 0
    For Each k In d.Keys
        parts(n) = Repr(k, depth - 1) & ": " & Repr(d(k), depth - 1)
        n = n + 1
    Next k
    repr_dictionary = Join(parts, ", ")
End Function

Private Function repr_array(ByRef arr As Variant, ByVal depth As Long) As String
    Dim nd As Long
    nd = array_dims(arr)
    Select Case nd
        Case 0:  repr_array = "Array()"
        Case 1:  repr_array = repr_1d(arr, depth)
        Case 2:  repr_array = repr_2d(arr, depth)
        Case Else: repr_array = array_summary(arr, nd)
    End Select
End Function

' 探测维度数；未定维空数组返回 0。对不存在的维度取 LBound 会抛错，
' 用 On Error 作"到达最高维"的终止信号（流控，非吞错）。
Private Function array_dims(ByRef arr As Variant) As Long
    Dim nd As Long
    Dim lo As Long, hi As Long
    On Error GoTo done
    nd = 0
    Do While True
        lo = LBound(arr, nd + 1)
        hi = UBound(arr, nd + 1)
        nd = nd + 1
    Loop
done:
    On Error GoTo 0
    array_dims = nd
End Function

Private Function repr_1d(ByRef arr As Variant, ByVal depth As Long) As String
    Dim lo As Long, hi As Long
    Dim i As Long
    Dim parts() As String
    lo = LBound(arr, 1)
    hi = UBound(arr, 1)
    If hi < lo Then
        repr_1d = "Array()"
        Exit Function
    End If
    ReDim parts(0 To hi - lo)
    For i = lo To hi
        parts(i - lo) = Repr(arr(i), depth - 1)
    Next i
    repr_1d = "[" & Join(parts, ", ") & "]"
End Function

Private Function repr_2d(ByRef arr As Variant, ByVal depth As Long) As String
    Dim r As Long, c As Long
    Dim rlo As Long, rhi As Long, clo As Long, chi As Long
    Dim rows() As String
    Dim cells() As String
    rlo = LBound(arr, 1): rhi = UBound(arr, 1)
    clo = LBound(arr, 2): chi = UBound(arr, 2)
    If rhi < rlo Or chi < clo Then
        repr_2d = "Array()"
        Exit Function
    End If
    ReDim rows(0 To rhi - rlo)
    For r = rlo To rhi
        ReDim cells(0 To chi - clo)
        For c = clo To chi
            cells(c - clo) = Repr(arr(r, c), depth - 1)
        Next c
        rows(r - rlo) = "[" & Join(cells, ", ") & "]"
    Next r
    repr_2d = "[" & Join(rows, ", ") & "]"
End Function

' 三维及以上数组逐元素展开会指数膨胀，输出结构摘要即可（维数 × 各维长度）
Private Function array_summary(ByRef arr As Variant, ByVal nd As Long) As String
    Dim i As Long
    Dim s As String
    s = ""
    For i = 1 To nd
        If i > 1 Then s = s & "x"
        s = s & CStr(UBound(arr, i) - LBound(arr, i) + 1)
    Next i
    array_summary = "Array<" & nd & "D " & s & ">"
End Function

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
