Attribute VB_Name = "mod_range"
'=====================================================================
' mod_range - Range 地址 通用工具（通用库 · v1.0）
'=====================================================================
' 职责：把 Range 转成便于记录/拼接的地址字符串。依赖 Excel 宿主，归入 xls/ 绑定层。
' 约定：rng Is Nothing 返回空串（不抛错）；absolute 控制是否含 $ 绝对引用。
'
' 目录 / Catalog
'   Range_Address(rng[, absolute])        As String   单元格绝对/相对地址
'   Range_SheetAddress(rng[, absolute])   As String   带工作表名的地址
'   Range_ShiftAddress(rng[, rows, cols]) As String   偏移后地址(带工作表名,不含$)
'   Range_JoinedAddress(rngs)             As String   多区域拼接地址
'=====================================================================
Option Explicit

' 单元格地址（默认绝对含 $）；rng 为空返回空串。
Public Function Range_Address(ByRef rng As Range, Optional ByVal absolute As Boolean = True) As String
    If rng Is Nothing Then Exit Function
    Range_Address = rng.Address(absolute)
End Function

' 带工作表名限定地址（'工作表名'!A1），工作表名含空格等时引号保险。
Public Function Range_SheetAddress(ByRef rng As Range, Optional ByVal absolute As Boolean = True) As String
    If rng Is Nothing Then Exit Function
    Range_SheetAddress = "'" & rng.Worksheet.Name & "'!" & rng.Address(absolute)
End Function

' 计算相对偏移 rows/cols 后的相对地址（不含 $，含工作表名），常用于"定位相邻单元格"。
Public Function Range_ShiftAddress(ByRef rng As Range, Optional ByVal rows As Long = 0, _
                                   Optional ByVal cols As Long = 0) As String
    Dim target As Range
    If rng Is Nothing Then Exit Function
    Set target = rng.Offset(rows, cols)
    Range_ShiftAddress = "'" & rng.Worksheet.Name & "'!" & target.Address(False, False)
End Function

' 把一组 Range（数组/集合）的地址用逗号拼接成多区域地址；自动跳过非 Range / Nothing 项。
Public Function Range_JoinedAddress(ByRef rngs As Variant) As String
    Dim r As Variant
    Dim parts As New Collection
    For Each r In rngs
        If Not r Is Nothing Then
            If TypeName(r) = "Range" Then parts.Add r.Address
        End If
    Next r
    Range_JoinedAddress = mod_string.String_Join(mod_array.Collection_To_Array(parts), ",")
End Function