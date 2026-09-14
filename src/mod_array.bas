Attribute VB_Name = "mod_array"
'=====================================================================
' mod_array - 数组 / 集合 通用工具（通用库 · v1.0）
'=====================================================================
' 命名体系：公共 API 使用 PascalCase + 域名前缀（Array_* / Collection_*），
'           内部过程与局部变量使用 snake_case（Python 风格）。
'
' 目录 / Catalog（新函数加入时在此登记）
'   Array_Contains(bytes as Variant 数组, val)         As Boolean   数组是否包含某元素
'   Array_Append(arr, elem)                            As Variant   尾部追加一个元素
'   Array_Extend(arr1, arr2)                           As Variant   拼接两个数组
'   Array_Distinct(arr)                                As Variant   去重（保持首见顺序）
'   Array_Map(arr, func_name)                          As Variant   高阶 map（逐元素回调）
'   Array_Filter(arr, func_name)                       As Variant   高阶 filter
'   Array_All(arr, func_name)                          As Boolean   高阶 every
'   Array_Any(arr, func_name)                          As Boolean   高阶 some
'   Collection_To_Array(col)                           As Variant   集合 -> 数组
'   Array_To_Collection(arr)                           As Collection 数组 -> 集合
'=====================================================================
Option Explicit

Public Function Array_Contains(ByRef arr As Variant, ByVal val As Variant) As Boolean
    Dim elem As Variant
    For Each elem In arr
        If elem = val Then
            Array_Contains = True
            Exit Function
        End If
    Next elem
    Array_Contains = False
End Function

Public Function Array_Append(ByRef arr As Variant, ByVal elem As Variant) As Variant
    Array_Append = Array_Extend(arr, Array(elem))
End Function

Public Function Array_Extend(ByRef arr1 As Variant, ByRef arr2 As Variant) As Variant
    Dim result() As Variant
    Dim i As Long, j As Long
    Dim n1 As Long, n2 As Long
    j = 0: n1 = 0: n2 = 0
    If IsArray(arr1) Then n1 = UBound(arr1) - LBound(arr1) + 1
    If IsArray(arr2) Then n2 = UBound(arr2) - LBound(arr2) + 1
    If n1 + n2 = 0 Then
        Array_Extend = Array()
        Exit Function
    End If
    ReDim result(0 To n1 + n2 - 1)
    If n1 > 0 Then
        For i = LBound(arr1) To UBound(arr1)
            result(j) = arr1(i): j = j + 1
        Next i
    End If
    If n2 > 0 Then
        For i = LBound(arr2) To UBound(arr2)
            result(j) = arr2(i): j = j + 1
        Next i
    End If
    Array_Extend = result
End Function

Public Function Array_Distinct(ByRef arr As Variant) As Variant
    Dim dict As Object
    Dim i As Long, idx As Long
    Dim k As Variant
    Dim unique() As Variant
    If Not IsArray(arr) Then
        Array_Distinct = Array()
        Exit Function
    End If
    Set dict = CreateObject("Scripting.Dictionary")
    For i = LBound(arr) To UBound(arr)
        If Not dict.Exists(arr(i)) Then dict.Add arr(i), True
    Next i
    If dict.Count = 0 Then
        Array_Distinct = Array()
        Exit Function
    End If
    ReDim unique(0 To dict.Count - 1)
    idx = 0
    For Each k In dict.Keys
        unique(idx) = k: idx = idx + 1
    Next k
    Array_Distinct = unique
End Function

'--- 高阶函数 ---------------------------------------------------------
' 回调 func_name 必须是宿主中可被 Application.Run 调用的过程/函数名。
' 原版用 On Error Resume Next 吞错；这里去掉以便尽早暴露回调错误。
Public Function Array_Map(ByRef arr As Variant, ByVal func_name As String) As Variant
    Dim result() As Variant
    Dim elem As Variant
    Dim i As Long
    If Not IsArray(arr) Then
        Array_Map = Array()
        Exit Function
    End If
    ReDim result(0 To UBound(arr) - LBound(arr))
    i = 0
    For Each elem In arr
        result(i) = Application.Run(func_name, elem)
        i = i + 1
    Next elem
    Array_Map = result
End Function

Public Function Array_Filter(ByRef arr As Variant, ByVal func_name As String) As Variant
    Dim ret As Collection
    Dim elem As Variant
    Set ret = New Collection
    If IsArray(arr) Then
        For Each elem In arr
            If Application.Run(func_name, elem) = True Then ret.Add elem
        Next elem
    End If
    If ret.Count = 0 Then
        Array_Filter = Array()
        Exit Function
    End If
    Array_Filter = Collection_To_Array(ret)
End Function

Public Function Array_All(ByRef arr As Variant, ByVal func_name As String) As Boolean
    Dim elem As Variant
    For Each elem In arr
        If Application.Run(func_name, elem) <> True Then
            Array_All = False
            Exit Function
        End If
    Next elem
    Array_All = True
End Function

Public Function Array_Any(ByRef arr As Variant, ByVal func_name As String) As Boolean
    Dim elem As Variant
    For Each elem In arr
        If Application.Run(func_name, elem) = True Then
            Array_Any = True
            Exit Function
        End If
    Next elem
    Array_Any = False
End Function

Public Function Collection_To_Array(ByRef col As Collection) As Variant
    Dim i As Long
    Dim arr() As Variant
    If col Is Nothing Or col.Count = 0 Then
        Collection_To_Array = Array()
        Exit Function
    End If
    ReDim arr(0 To col.Count - 1)
    For i = 0 To col.Count - 1
        arr(i) = col.Item(i + 1)
    Next i
    Collection_To_Array = arr
End Function

Public Function Array_To_Collection(ByRef arr As Variant) As Collection
    Dim elem As Variant
    Dim clc As New Collection
    Dim idx As Long
    idx = 1
    For Each elem In arr
        clc.Add elem, CStr(idx)
        idx = idx + 1
    Next elem
    Set Array_To_Collection = clc
End Function