Attribute VB_Name = "mod_debug"
'=====================================================================
' mod_debug - 开发期调试打印工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   Show_Arr_Members(arr)   打印数组维度信息(类型/上下界/元素数)
'   Print_Array(arr)        打印数组元素(逗号分隔)
'   Print_Clc(col)          打印集合元素
'   Print_Dict(dict)        打印字典键值
'   Print_Lines(lines)      逐行打印字符串数组
' 说明：本模块仅输出到 VBE 立即窗口(Debug.Print)，用于开发期排查，
'       不作为运行期结果来源（运行期断言请用 mod_tests）。
'=====================================================================
Option Explicit

Public Sub Show_Arr_Members(ByRef arr As Variant)
    Dim nd As Long
    Dim lo As Long, hi As Long
    If Not IsArray(arr) Then
        Debug.Print "not an array, type=" & TypeName(arr)
        Exit Sub
    End If
    ' 逐维探测维度数与上下界：对不存在维度调用 LBound/UBound 会抛错，
    ' 用 On Error 作“到达最高维”的终止信号（这里是流控，非吞错）。
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