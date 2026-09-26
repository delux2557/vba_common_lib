Attribute VB_Name = "mod_range"
'=====================================================================
' mod_range - Range 地址 / 批量读写 通用工具（通用库 · v1.1）
'=====================================================================
' 职责：把 Range 转成便于记录/拼接的地址字符串，以及一次读/写整块区域的
'       批量操作（数组往返，避免逐格读写——性能差一个数量级）。
'       依赖 Excel 宿主，归入 xls/ 绑定层。
' 约定：rng Is Nothing 时返回类函数返回空串（不抛错）；absolute 控制是否含 $ 绝对引用。
'
' 目录 / Catalog
'   Range_Address(rng[, absolute])        As String   单元格绝对/相对地址
'   Range_SheetAddress(rng[, absolute])   As String   带工作表名的地址
'   Range_ShiftAddress(rng[, rows, cols]) As String   偏移后地址(带工作表名,不含$)
'   Range_JoinedAddress(rngs)             As String   多区域拼接地址
'   Range_ReadValues(rng)                 As Variant  一次读整块区域(原生 Value 语义)
'   Range_WriteValues(rng, values[, as_column])        一次写整块(数组自动 Resize/标量填充,防闪烁)
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

'--- 批量读写（数组往返，替代逐格循环；性能差一个数量级） -----------------

' 一次读整块区域，返回原生 Range.Value 语义：
'   单格 → 标量；整行/整列 → 一维数组；矩形多格 → 二维数组(1 基, 行, 列)。
' rng 为 Nothing 返回 Empty；多区域抛错（避免隐藏行为）。
Public Function Range_ReadValues(ByRef rng As Range) As Variant
    If rng Is Nothing Then Exit Function
    If rng.Areas.Count > 1 Then _
        Err.Raise 45000, "mod_range", "Range_ReadValues: 不支持多区域，请按区域分别读取"
    Range_ReadValues = rng.Value
End Function

' 一次写整块区域。values 形态决定行为：
'   二维数组 → 以 rng 左上角为锚点 Resize 到数组尺寸后一次赋值；
'   一维数组 → 默认按行展开；as_column=True 时转置按列展开；
'   标量     → 填充整个 rng 区域。
' 写入期间临时关闭 ScreenUpdating 防闪烁，成败均恢复（副作用安全）；
' 多区域 / Nothing / 空数组 / 对象等非法输入抛错 #45000。
Public Sub Range_WriteValues(ByRef rng As Range, ByRef values As Variant, _
                             Optional ByVal as_column As Boolean = False)
    Dim saved_ui As Boolean
    If rng Is Nothing Then Err.Raise 45000, "mod_range", "Range_WriteValues: rng 为 Nothing"
    If rng.Areas.Count > 1 Then _
        Err.Raise 45000, "mod_range", "Range_WriteValues: 不支持多区域，请按区域分别写入"
    If IsObject(values) Or IsEmpty(values) Or IsNull(values) Then _
        Err.Raise 45000, "mod_range", "Range_WriteValues: values 须为标量或一/二维数组"

    saved_ui = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo restore
    If IsArray(values) Then
        write_array rng, values, as_column
    Else
        rng.Value = values
    End If
restore:
    Application.ScreenUpdating = saved_ui
    If Err.Number <> 0 Then Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' 按数组维度分发写入；一维数组用 Application.Transpose 转置成列（>65536 元素受限）。
Private Sub write_array(ByRef rng As Range, ByRef arr As Variant, ByVal as_column As Boolean)
    Dim n As Long
    Dim rows As Long, cols As Long
    n = UBound(arr, 1)
    If n < LBound(arr, 1) Then Err.Raise 45000, "mod_range", "Range_WriteValues: 空数组不可写"
    On Error GoTo one_dim
    cols = UBound(arr, 2) - LBound(arr, 2) + 1
    On Error GoTo 0
    rows = n - LBound(arr, 1) + 1
    rng.Resize(rows, cols).Value = arr
    Exit Sub
one_dim:
    On Error GoTo 0
    rows = n - LBound(arr, 1) + 1
    If as_column Then
        rng.Resize(rows, 1).Value = Application.Transpose(arr)
    Else
        rng.Resize(1, rows).Value = arr
    End If
End Sub