Attribute VB_Name = "mod_dict"
'=====================================================================
' mod_dict - 字典 通用工具（通用库 · v1.1）
'=====================================================================
' 目录 / Catalog
'   Dict_Create()                     As Object  新建字典(晚绑定 Scripting.Dictionary)
'   Dict_Set(dict, key, val)                    写入/覆盖键值
'   Dict_Get(dict, key[, default])    As Variant  读值，缺键返回 default
'   Dict_Exists(dict, key)            As Boolean  键是否存在
'   Dict_Count(dict)                  As Long     键数量
'   Dict_Keys(dict)                   As Variant  所有键(数组)
'   Dict_Values(dict)                 As Variant  所有值(数组)
'   Dict_To_Array(dict)               As Variant  -> 二维数组(行=条目, 列1=键,列2=值)
'   Dict_FromArray(pairs)             As Object   二维数组(键,值) -> 字典(Dict_To_Array 逆)
'   Dict_Remove(dict, key)                       移除键
'   Dict_Clear(dict)                            清空
' 说明：晚绑定 Scripting.Dictionary，无需勾选外部引用；键区分大小写。
'=====================================================================
Option Explicit

Public Function Dict_Create() As Object
    Set Dict_Create = CreateObject("Scripting.Dictionary")
End Function

Public Sub Dict_Set(ByRef dict As Object, ByVal key As Variant, ByVal val As Variant)
    If dict.Exists(key) Then dict.Remove key
    dict.Add key, val
End Sub

Public Function Dict_Get(ByRef dict As Object, ByVal key As Variant, Optional ByVal default As Variant = vbNullString) As Variant
    If dict Is Nothing Or Not dict.Exists(key) Then
        Dict_Get = default
    Else
        Dict_Get = dict(key)
    End If
End Function

Public Function Dict_Exists(ByRef dict As Object, ByVal key As Variant) As Boolean
    Dict_Exists = (Not dict Is Nothing) And dict.Exists(key)
End Function

Public Function Dict_Count(ByRef dict As Object) As Long
    If dict Is Nothing Then
        Dict_Count = 0
    Else
        Dict_Count = dict.Count
    End If
End Function

Public Function Dict_Keys(ByRef dict As Object) As Variant
    Dim k As Variant
    Dim n As Long
    Dim i As Long
    Dim arr() As Variant
    If dict Is Nothing Then
        Dict_Keys = Array()
        Exit Function
    End If
    n = dict.Count
    If n = 0 Then
        Dict_Keys = Array()
        Exit Function
    End If
    ReDim arr(0 To n - 1)
    i = 0
    For Each k In dict.Keys
        arr(i) = k: i = i + 1
    Next k
    Dict_Keys = arr
End Function

Public Function Dict_Values(ByRef dict As Object) As Variant
    Dim k As Variant
    Dim n As Long
    Dim i As Long
    Dim arr() As Variant
    If dict Is Nothing Then
        Dict_Values = Array()
        Exit Function
    End If
    n = dict.Count
    If n = 0 Then
        Dict_Values = Array()
        Exit Function
    End If
    ReDim arr(0 To n - 1)
    i = 0
    For Each k In dict.Keys
        arr(i) = dict(k): i = i + 1
    Next k
    Dict_Values = arr
End Function

' 二维数组 1 基：(行, 1=键, 2=值)，便于直接写出到单元格。
Public Function Dict_To_Array(ByRef dict As Object) As Variant
    Dim k As Variant
    Dim n As Long
    Dim i As Long
    Dim arr() As Variant
    If dict Is Nothing Or dict.Count = 0 Then
        Dict_To_Array = Array()
        Exit Function
    End If
    n = dict.Count
    ReDim arr(1 To n, 1 To 2)
    i = 1
    For Each k In dict.Keys
        arr(i, 1) = k
        arr(i, 2) = dict(k)
        i = i + 1
    Next k
    Dict_To_Array = arr
End Function

Public Function Dict_FromArray(ByRef pairs As Variant) As Object
    Dim d As Object
    Dim r As Long, r0 As Long, r1 As Long
    Set d = CreateObject("Scripting.Dictionary")
    If IsArray(pairs) Then
        r0 = LBound(pairs, 1): r1 = UBound(pairs, 1)
        For r = r0 To r1
            d.Add pairs(r, 1), pairs(r, 2)
        Next r
    End If
    Set Dict_FromArray = d
End Function

Public Sub Dict_Remove(ByRef dict As Object, ByVal key As Variant)
    If Not dict Is Nothing Then
        If dict.Exists(key) Then dict.Remove key
    End If
End Sub

Public Sub Dict_Clear(ByRef dict As Object)
    If Not dict Is Nothing Then dict.RemoveAll
End Sub