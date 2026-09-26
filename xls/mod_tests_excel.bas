Attribute VB_Name = "mod_tests_excel"
'=====================================================================
' mod_tests_excel - Excel 绑定层单元测试（xls/ · 自包含脚手架）
'=====================================================================
' 用法：
'   fail_count = mod_tests_excel.Run_All_Tests(report_path)
'   返回值为失败数；report_path 为空则仅输出立即窗口。
' 说明：本模块测试 xls/ 层（mod_workbook / mod_range），位于 Excel 宿主内
'       运行，因此必然依赖 ThisWorkbook/Range/Application。纯净的 src/ 层
'       测试请用 src/mod_tests。
'=====================================================================
Option Explicit

Private p_pass As Long
Private p_fail As Long
Private p_report As String
Private p_last_topic As String

Private Sub log_line(ByVal s As String)
    Debug.Print s
    p_report = p_report & s & vbCrLf
End Sub

Private Sub Test_Equal(ByVal actual As Variant, ByVal expected As Variant, ByVal label As String)
    p_last_topic = label
    If CStr(actual) = CStr(expected) Then
        p_pass = p_pass + 1
        log_line "[PASS] " & label
    Else
        p_fail = p_fail + 1
        log_line "[FAIL] " & label & " | expected=[" & CStr(expected) & "] actual=[" & CStr(actual) & "]"
    End If
End Sub

Private Sub Test_True(ByVal cond As Boolean, ByVal label As String)
    p_last_topic = label
    If cond Then p_pass = p_pass + 1 Else p_fail = p_fail + 1
    log_line IIf(cond, "[PASS] ", "[FAIL] ") & label
End Sub

Private Sub Test_False(ByVal cond As Boolean, ByVal label As String)
    Test_True (Not cond), label
End Sub

'--- 运行入口 -------------------------------------------------------
Public Function Run_All_Tests(Optional ByVal report_path As String = vbNullString) As Long
    p_pass = 0: p_fail = 0: p_report = ""
    log_line "===== VBA Common Lib (Excel 绑定层) 单元测试开始 ====="
    run_protected "suite_workbook"
    run_protected "suite_range"
    run_protected "suite_demo"
    log_line "===== 结束：通过 " & p_pass & " / 失败 " & p_fail & " ====="
    If Len(report_path) > 0 Then mod_file.File_Write report_path, p_report, False, True
    Run_All_Tests = p_fail
End Function

Private Sub run_protected(ByVal suite As String)
    p_last_topic = suite & " (套件入口)"
    On Error GoTo crash
    Select Case suite
        Case "suite_workbook": suite_workbook
        Case "suite_range":    suite_range
        Case "suite_demo":     suite_demo
    End Select
    Exit Sub
crash:
    p_fail = p_fail + 1
    log_line "[CRASH] " & suite & " 抛错 #" & Err.Number & " in " & Err.Source & ": " & Err.Description & _
             " ~ 最近步骤: " & p_last_topic
End Sub

'--- suites ---------------------------------------------------------
Private Sub suite_workbook()
    log_line "--- mod_workbook ---"
    Test_True mod_workbook.Sheet_Exists("Sheet1", ThisWorkbook), "Sheet_Exists: Sheet1 存在"
    Test_False mod_workbook.Sheet_Exists("no_such_sheet_9", ThisWorkbook), "Sheet_Exists: 不存在"
    Test_True mod_workbook.Workbook_Exists(ThisWorkbook.Name), "Workbook_Exists: 自身"
    Dim names As Variant
    names = mod_workbook.Workbook_SheetNames(ThisWorkbook)
    Test_True mod_array.Array_Contains(names, "Sheet1"), "Workbook_SheetNames: 含 Sheet1"

    '--- 工作表批量管理（临时表统一 T_ 前缀，套件末尾兜底清理） ---------
    Test_True (mod_workbook.Workbook_SheetCount(ThisWorkbook) > 0), "Workbook_SheetCount: >0"
    Test_True (Not mod_workbook.Sheet_Get("Sheet1", ThisWorkbook) Is Nothing), "Sheet_Get: 存在返回对象"
    Test_True (mod_workbook.Sheet_Get("no_such_sheet_9", ThisWorkbook) Is Nothing), "Sheet_Get: 不存在返回 Nothing"

    Dim ws_t As Worksheet
    Set ws_t = mod_workbook.Sheet_Ensure("T_ws_ensure", ThisWorkbook)
    Test_True (Not ws_t Is Nothing), "Sheet_Ensure: 不存在则新建"
    Test_True mod_workbook.Sheet_Exists("T_ws_ensure", ThisWorkbook), "Sheet_Ensure: 新建后存在"
    Dim ws_t2 As Worksheet
    Set ws_t2 = mod_workbook.Sheet_Ensure("T_ws_ensure", ThisWorkbook)
    Test_True (ws_t2 Is ws_t), "Sheet_Ensure: 已存在幂等返回同一对象"

    Dim caught As Boolean
    caught = False
    On Error Resume Next
    mod_workbook.Sheet_Create "T_ws_dup", ThisWorkbook
    mod_workbook.Sheet_Create "T_ws_dup", ThisWorkbook
    If Err.Number <> 0 Then caught = True
    On Error GoTo 0
    Test_True caught, "Sheet_Create: 重名抛错"
    mod_workbook.Sheet_Delete mod_workbook.Sheet_Get("T_ws_dup", ThisWorkbook)

    Dim act_before As Worksheet
    Set act_before = ThisWorkbook.ActiveSheet
    Dim ws_copy As Worksheet
    Set ws_copy = mod_workbook.Sheet_Copy(ws_t)
    Test_True (Not ws_copy Is Nothing), "Sheet_Copy: 返回副本对象"
    Test_True mod_workbook.Sheet_Exists("T_ws_ensure (2)", ThisWorkbook), "Sheet_Copy: 默认命名 原名 (2)"
    Test_True (ThisWorkbook.ActiveSheet Is act_before), "Sheet_Copy: 复制后恢复活动表"

    mod_workbook.Sheet_Rename ws_copy, "T_ws_copied"
    Test_False mod_workbook.Sheet_Exists("T_ws_ensure (2)", ThisWorkbook), "Sheet_Rename: 旧名不再存在"
    Test_True mod_workbook.Sheet_Exists("T_ws_copied", ThisWorkbook), "Sheet_Rename: 新名存在"
    caught = False
    On Error Resume Next
    mod_workbook.Sheet_Rename ws_copy, "T_ws_ensure"
    If Err.Number <> 0 Then caught = True
    On Error GoTo 0
    Test_True caught, "Sheet_Rename: 重名抛错"

    mod_workbook.Sheet_Delete ws_copy
    Test_False mod_workbook.Sheet_Exists("T_ws_copied", ThisWorkbook), "Sheet_Delete: 删除后不存在"
    mod_workbook.Sheet_Delete ws_t
    Test_False mod_workbook.Sheet_Exists("T_ws_ensure", ThisWorkbook), "Sheet_Delete: 删除源表"

    ' 兜底清理：删除所有 T_ 前缀临时表（逐次遍历避免 For Each 删除跳项）
    Dim sh_clean As Worksheet
    Dim found_t As Boolean
    Do
        found_t = False
        For Each sh_clean In ThisWorkbook.Sheets
            If Left$(sh_clean.Name, 2) = "T_" Then
                mod_workbook.Sheet_Delete sh_clean
                found_t = True
                Exit For
            End If
        Next sh_clean
    Loop While found_t
End Sub

Private Sub suite_range()
    log_line "--- mod_range ---"
    Dim r As Range
    Set r = ThisWorkbook.Sheets(1).Range("A1")
    Test_Equal mod_range.Range_Address(r), "$A$1", "Range_Address"
    Test_Equal mod_range.Range_ShiftAddress(r, 1, 0), "'" & r.Worksheet.Name & "'!A2", "Range_ShiftAddress 下移一行"
    Test_True (InStr(1, mod_range.Range_SheetAddress(r), "'" & r.Worksheet.Name & "'!") > 0), "Range_SheetAddress: 含工作表名"

    ' Repr 的 Range 分支（纯层 Repr 分发到 Excel Range 类型）
    Set r = ThisWorkbook.Sheets(1).Range("A1")
    r.Value = "hello"
    Test_Equal mod_debug.Repr(r), "Range(""A1"", Value=""hello"")", "Repr: Range(地址+值)"
    Set r = ThisWorkbook.Sheets(1).Range("A1:B2")
    r.Value = 5
    Test_Equal mod_debug.Repr(r), "Range(""A1:B2"", Value=[[5, 5], [5, 5]])", "Repr: 多单元格 Range 值数组"

    ' Range_WriteValues / Range_ReadValues 批量读写（D 列起，避开 A 列既有断言）
    Dim r_block As Range
    Set r_block = ThisWorkbook.Sheets(1).Range("D1:E2")
    Dim src2d(1 To 2, 1 To 2) As Variant
    src2d(1, 1) = 10: src2d(1, 2) = 20
    src2d(2, 1) = 30: src2d(2, 2) = 40
    mod_range.Range_WriteValues r_block, src2d
    Dim got As Variant
    got = mod_range.Range_ReadValues(r_block)
    Test_Equal CStr(got(1, 1)) & "|" & CStr(got(1, 2)) & "|" & CStr(got(2, 1)) & "|" & CStr(got(2, 2)), _
                "10|20|30|40", "Range_Write/ReadValues: 二维数组往返"

    mod_range.Range_WriteValues r_block, 7
    got = mod_range.Range_ReadValues(r_block)
    Test_Equal CStr(got(1, 1)) & "|" & CStr(got(2, 2)), "7|7", "Range_WriteValues: 标量填充全区"

    Dim r_cell As Range
    Set r_cell = ThisWorkbook.Sheets(1).Range("D4")
    Dim one_d As Variant
    one_d = Array(1, 2, 3)
    mod_range.Range_WriteValues r_cell, one_d
    Test_Equal CStr(r_cell.Value) & "|" & CStr(r_cell.Offset(0, 1).Value) & "|" & CStr(r_cell.Offset(0, 2).Value), _
                "1|2|3", "Range_WriteValues: 一维数组按行展开"

    Set r_cell = ThisWorkbook.Sheets(1).Range("D5")
    mod_range.Range_WriteValues r_cell, one_d, True
    Test_Equal CStr(r_cell.Value) & "|" & CStr(r_cell.Offset(1, 0).Value) & "|" & CStr(r_cell.Offset(2, 0).Value), _
                "1|2|3", "Range_WriteValues: 一维数组按列(as_column)"

    ' 非法输入抛错
    Dim r_anchor As Range
    Set r_anchor = ThisWorkbook.Sheets(1).Range("A1")
    Dim empty_arr As Variant
    empty_arr = Array()
    Dim caught2 As Boolean
    caught2 = False
    On Error Resume Next
    mod_range.Range_WriteValues r_anchor, empty_arr
    If Err.Number <> 0 Then caught2 = True
    Err.Clear
    On Error GoTo 0
    Test_True caught2, "Range_WriteValues: 空数组抛错"

    Dim r_multi As Range
    Set r_multi = ThisWorkbook.Sheets(1).Range("A1,B2")
    caught2 = False
    On Error Resume Next
    mod_range.Range_WriteValues r_multi, 1
    If Err.Number <> 0 Then caught2 = True
    Err.Clear
    On Error GoTo 0
    Test_True caught2, "Range_WriteValues: 多区域抛错"

    caught2 = False
    On Error Resume Next
    mod_range.Range_ReadValues r_multi
    If Err.Number <> 0 Then caught2 = True
    Err.Clear
    On Error GoTo 0
    Test_True caught2, "Range_ReadValues: 多区域抛错"
End Sub

Private Sub suite_demo()
    log_line "--- mod_demo ---"
    ' 运行完整演示：造数据(60x2, 噪声±0.8, 真系数[2,3,-1], seed=7) → 训练 → 预测 → 评估 → 写表
    mod_demo.Demo_Ml_LinearReg
    Dim ws_d As Worksheet
    Set ws_d = ThisWorkbook.Worksheets("线性回归演示")
    Test_True (Not ws_d Is Nothing), "Demo: 生成演示工作表"
    If ws_d Is Nothing Then Exit Sub
    ' 数据表：表头 + 60 行数据（A1 起）
    Test_Equal CStr(ws_d.Cells(1, 1).Value), "特征1", "Demo: 数据表表头"
    Test_Equal CStr(ws_d.Cells(1, 3).Value), "y(标签)", "Demo: 数据表标签列"
    Test_True (ws_d.Cells(61, 3).Value <> vbNullString), "Demo: 60 行数据已写入"
    ' 系数区：按标签文本 Find 定位（布局变动不破坏断言）
    Dim r_tag As Range
    Set r_tag = ws_d.Columns(1).Find("截距 b0", LookAt:=xlWhole)
    Test_True (Not r_tag Is Nothing), "Demo: 找到截距标签"
    If r_tag Is Nothing Then Exit Sub
    Test_True (Abs(r_tag.Offset(0, 1).Value - 2#) < 0.5), "Demo: 截距逼近 2"
    Set r_tag = ws_d.Columns(1).Find("权重 w1", LookAt:=xlWhole)
    Test_True (Not r_tag Is Nothing), "Demo: 找到 w1 标签"
    If r_tag Is Nothing Then Exit Sub
    Test_True (Abs(r_tag.Offset(0, 1).Value - 3#) < 0.3), "Demo: w1 逼近 3"
    Set r_tag = ws_d.Columns(1).Find("权重 w2", LookAt:=xlWhole)
    Test_True (Not r_tag Is Nothing), "Demo: 找到 w2 标签"
    If r_tag Is Nothing Then Exit Sub
    Test_True (Abs(r_tag.Offset(0, 1).Value + 1#) < 0.3), "Demo: w2 逼近 -1"
    Set r_tag = ws_d.Columns(1).Find("拟合优度 R2", LookAt:=xlWhole)
    Test_True (Not r_tag Is Nothing), "Demo: 找到 R2 标签"
    If r_tag Is Nothing Then Exit Sub
    Test_True (r_tag.Offset(0, 1).Value > 0.9), "Demo: R2 高"
    ' 预测对比：按"模型预测"表头定位列，其下 3 行依次是 3 个新样本的预测
    Dim r_mp As Range
    Set r_mp = ws_d.Cells.Find(What:="模型预测", LookAt:=xlWhole)
    Test_True (Not r_mp Is Nothing), "Demo: 找到预测表"
    If r_mp Is Nothing Then Exit Sub
    Test_True (Abs(r_mp.Offset(1, 0).Value - 4#) < 0.5), "Demo: 预测(1.5,2.5) 逼近 4"
    Test_True (Abs(r_mp.Offset(2, 0).Value - 11#) < 0.5), "Demo: 预测(4,3) 逼近 11"
    Test_True (Abs(r_mp.Offset(3, 0).Value - 22.5) < 0.5), "Demo: 预测(7,0.5) 逼近 22.5"
    ' 清理演示表
    mod_workbook.Sheet_Delete ws_d
    Test_False mod_workbook.Sheet_Exists("线性回归演示", ThisWorkbook), "Demo: 清理演示表"
End Sub