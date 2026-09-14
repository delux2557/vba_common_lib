Attribute VB_Name = "mod_sort"
'=====================================================================
' mod_sort - 排序 / 反转 通用工具（通用库 · v1.1）
'=====================================================================
' 目录 / Catalog
'   Array_Sort(arr[, ascending[, key_func]])   As Variant  稳定排序(返回新数组)
'   Array_Reverse(arr)                         As Variant  反转(返回新数组)
' 说明：
'   - 不改动入参，返回排序/反转后的新数组（0 基，与原数组元素数一致）。
'   - 元素需同类型可比（数值 / 字符串 / 日期），请勿混用不同类型。
'   - 稳定排序：排序键相等的元素保持原始相对顺序。
'   - key_func 可选：经 Application.Run 对每个元素计算排序键（可提取列 / 属性）。
'=====================================================================
Option Explicit

Public Function Array_Reverse(ByRef arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim new_arr() As Variant
    If Not IsArray(arr) Then
        Array_Reverse = Array()
        Exit Function
    End If
    n = UBound(arr) - LBound(arr) + 1
    If n < 1 Then
        Array_Reverse = Array()
        Exit Function
    End If
    ReDim new_arr(0 To n - 1)
    For i = 0 To n - 1
        new_arr(i) = arr(LBound(arr) + (n - 1 - i))
    Next i
    Array_Reverse = new_arr
End Function

Public Function Array_Sort(ByRef arr As Variant, Optional ByVal ascending As Boolean = True, _
                           Optional ByVal key_func As String = vbNullString) As Variant
    Dim n As Long
    Dim i As Long
    Dim keys() As Variant
    Dim vals() As Variant
    Dim idx() As Long
    Dim tmp() As Long
    Dim out_arr() As Variant

    If Not IsArray(arr) Then
        Array_Sort = Array()
        Exit Function
    End If
    n = UBound(arr) - LBound(arr) + 1
    If n < 1 Then
        Array_Sort = Array()
        Exit Function
    End If
    If n <= 1 Then
        Array_Sort = arr_0based(arr)
        Exit Function
    End If

    ' 预计算排序键（可选回调），同时保存原值
    ReDim keys(0 To n - 1)
    ReDim vals(0 To n - 1)
    ReDim idx(0 To n - 1)
    For i = 0 To n - 1
        vals(i) = arr(LBound(arr) + i)
        If Len(key_func) > 0 Then
            keys(i) = Application.Run(key_func, vals(i))
        Else
            keys(i) = vals(i)
        End If
        idx(i) = i
    Next i

    ReDim tmp(0 To n - 1)
    MergeSort keys, idx, tmp, 0, n - 1, ascending

    ReDim out_arr(0 To n - 1)
    For i = 0 To n - 1
        out_arr(i) = vals(idx(i))
    Next i
    Array_Sort = out_arr
End Function

Private Function arr_0based(ByRef arr As Variant) As Variant
    Dim n As Long
    Dim i As Long
    Dim c() As Variant
    n = UBound(arr) - LBound(arr) + 1
    If n < 1 Then
        arr_0based = Array()
        Exit Function
    End If
    ReDim c(0 To n - 1)
    For i = 0 To n - 1
        c(i) = arr(LBound(arr) + i)
    Next i
    arr_0based = c
End Function

Private Sub MergeSort(ByRef keys() As Variant, ByRef idx() As Long, ByRef tmp() As Long, _
                      ByVal lo As Long, ByVal hi As Long, ByVal ascending As Boolean)
    Dim mid As Long
    If lo >= hi Then Exit Sub
    mid = (lo + hi) \ 2
    MergeSort keys, idx, tmp, lo, mid, ascending
    MergeSort keys, idx, tmp, mid + 1, hi, ascending
    Merge keys, idx, tmp, lo, mid, hi, ascending
End Sub

Private Sub Merge(ByRef keys() As Variant, ByRef idx() As Long, ByRef tmp() As Long, _
                  ByVal lo As Long, ByVal mid As Long, ByVal hi As Long, ByVal ascending As Boolean)
    Dim i As Long, j As Long, k As Long
    For k = lo To hi: tmp(k) = idx(k): Next k
    i = lo: j = mid + 1: k = lo
    Do While i <= mid And j <= hi
        If take_left(keys(tmp(i)), keys(tmp(j)), ascending) Then
            idx(k) = tmp(i): i = i + 1
        Else
            idx(k) = tmp(j): j = j + 1
        End If
        k = k + 1
    Loop
    Do While i <= mid
        idx(k) = tmp(i): i = i + 1: k = k + 1
    Loop
    Do While j <= hi
        idx(k) = tmp(j): j = j + 1: k = k + 1
    Loop
End Sub

' 决定合并时是否取左侧：升序取 lk<=rk，降序取 lk>=rk；相等恒取左侧以保证稳定。
Private Function take_left(ByVal lk As Variant, ByVal rk As Variant, ByVal ascending As Boolean) As Boolean
    If ascending Then
        take_left = Not (lk > rk)
    Else
        take_left = Not (lk < rk)
    End If
End Function