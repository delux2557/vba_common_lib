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
End Sub

Private Sub suite_range()
    log_line "--- mod_range ---"
    Dim r As Range
    Set r = ThisWorkbook.Sheets(1).Range("A1")
    Test_Equal mod_range.Range_Address(r), "$A$1", "Range_Address"
    Test_Equal mod_range.Range_ShiftAddress(r, 1, 0), "'" & r.Worksheet.Name & "'!A2", "Range_ShiftAddress 下移一行"
    Test_True (InStr(1, mod_range.Range_SheetAddress(r), "'" & r.Worksheet.Name & "'!") > 0), "Range_SheetAddress: 含工作表名"
End Sub