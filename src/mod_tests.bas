Attribute VB_Name = "mod_tests"
'=====================================================================
' mod_tests - 通用库单元测试（TDD 文化：新函数必须在此登记一条断言）
'=====================================================================
' 用法：
'   fail_count = mod_tests.Run_All_Tests(report_path)
'   返回值为失败数；report_path 为空则仅输出立即窗口。
' 也可按需只跑单独 suite：suite_array / suite_string / ...
'=====================================================================
Option Explicit

Private p_pass As Long
Private p_fail As Long
Private p_report As String
Private p_last_topic As String   ' 最近一条断言标签，供运行时错误定位

'--- 断言工具 -------------------------------------------------------
Private Sub log_line(ByVal s As String)
    Debug.Print s
    p_report = p_report & s & vbCrLf
End Sub

Public Sub Test_Equal(ByVal actual As Variant, ByVal expected As Variant, ByVal label As String)
    p_last_topic = label
    If CStr(actual) = CStr(expected) Then
        p_pass = p_pass + 1
        log_line "[PASS] " & label
    Else
        p_fail = p_fail + 1
        log_line "[FAIL] " & label & " | expected=[" & CStr(expected) & "] actual=[" & CStr(actual) & "]"
    End If
End Sub

Public Sub Test_True(ByVal cond As Boolean, ByVal label As String)
    p_last_topic = label
    If cond Then p_pass = p_pass + 1 Else p_fail = p_fail + 1
    log_line IIf(cond, "[PASS] ", "[FAIL] ") & label
End Sub

Public Sub Test_False(ByVal cond As Boolean, ByVal label As String)
    Test_True (Not cond), label
End Sub

' 数组规约成规范字符串便于断言
Public Function canon(ByRef arr As Variant) As String
    canon = mod_string.String_Join(arr, ",")
End Function

Private Function test_tmp_dir() As String
    Dim d As String
    d = Environ("TEMP") & "\vba_common_test_" & mod_date.Date_Stamp("time")
    mod_file.Folder_Ensure d
    test_tmp_dir = d
End Function

'--- 高阶函数回调（Public，供 Application.Run 调用） -----------------
Public Function helper_is_even(ByVal num As Variant) As Boolean
    helper_is_even = (num Mod 2 = 0)
End Function

Public Function helper_mul_100(ByVal num As Variant) As Variant
    helper_mul_100 = num * 100
End Function

Public Function helper_key_mod10(ByVal num As Variant) As Variant
    helper_key_mod10 = num Mod 10
End Function

'--- 运行入口 -------------------------------------------------------
Public Function Run_All_Tests(Optional ByVal report_path As String = vbNullString) As Long
    p_pass = 0: p_fail = 0: p_report = ""
    log_line "===== VBA Common Lib 单元测试开始 ====="
    run_protected "suite_array"
    run_protected "suite_string"
    run_protected "suite_regex"
    run_protected "suite_date"
    run_protected "suite_file"
    run_protected "suite_dict"
    run_protected "suite_sort"
    run_protected "suite_json"
    run_protected "suite_version"
    log_line "===== 结束：通过 " & p_pass & " / 失败 " & p_fail & " ====="
    If Len(report_path) > 0 Then mod_file.File_Write report_path, p_report, False, True
    Run_All_Tests = p_fail
End Function

' 逐个执行 suite；任一 suite 抛运行时错误时记录报错节点并继续，不中断整体。
Private Sub run_protected(ByVal suite As String)
    p_last_topic = suite & " (套件入口)"
    On Error GoTo crash
    Select Case suite
        Case "suite_array":    suite_array
        Case "suite_string":   suite_string
        Case "suite_regex":    suite_regex
        Case "suite_date":     suite_date
        Case "suite_file":     suite_file
        Case "suite_dict":     suite_dict
        Case "suite_sort":     suite_sort
        Case "suite_json":     suite_json
        Case "suite_version":  suite_version
    End Select
    Exit Sub
crash:
    p_fail = p_fail + 1
    log_line "[CRASH] " & suite & " 抛错 #" & Err.Number & " in " & Err.Source & ": " & Err.Description & _
             " ~ 最近步骤: " & p_last_topic
End Sub

'--- suites ---------------------------------------------------------
Private Sub suite_array()
    log_line "--- mod_array ---"
    Test_True mod_array.Array_Contains(Array(11, 22, 33), 22), "Array_Contains: 命中"
    Test_False mod_array.Array_Contains(Array(11, 22, 33), 99), "Array_Contains: 未命中"
    Test_Equal canon(mod_array.Array_Append(Array(1, 2), 3)), "1,2,3", "Array_Append"
    Test_Equal canon(mod_array.Array_Extend(Array(1, 2), Array(3, 4))), "1,2,3,4", "Array_Extend"
    Test_Equal canon(mod_array.Array_Extend(Array(), Array(3, 4))), "3,4", "Array_Extend: 空数组拼接"
    Test_Equal canon(mod_array.Array_Distinct(Array(1, 2, 2, 3, 1, "a", "a"))), "1,2,3,a", "Array_Distinct"
    Test_Equal canon(mod_array.Array_Map(Array(1, 2, 3), "helper_mul_100")), "100,200,300", "Array_Map"
    Test_Equal canon(mod_array.Array_Filter(Array(1, 2, 3, 4), "helper_is_even")), "2,4", "Array_Filter"
    Test_True mod_array.Array_All(Array(2, 4, 6), "helper_is_even"), "Array_All: all even"
    Test_False mod_array.Array_All(Array(2, 3, 4), "helper_is_even"), "Array_All: not all even"
    Test_False mod_array.Array_Any(Array(1, 3, 5), "helper_is_even"), "Array_Any: none even"
    Test_True mod_array.Array_Any(Array(1, 2, 5), "helper_is_even"), "Array_Any: some even"
    Dim clc As Collection
    Set clc = mod_array.Array_To_Collection(Array("a", "b", "c"))
    Test_Equal canon(mod_array.Collection_To_Array(clc)), "a,b,c", "Collection 往返转换"
End Sub

Private Sub suite_string()
    log_line "--- mod_string ---"
    Test_Equal mod_string.String_Trim("  abc  "), "abc", "String_Trim: 默认去空白"
    Test_Equal mod_string.String_Trim(vbTab & "x" & vbNewLine), "x", "String_Trim: 去制表换行"
    Test_Equal mod_string.String_Trim("xxabcxx", "x"), "abc", "String_Trim: 指定字符集"
    Test_Equal mod_string.String_LTrim("   abc"), "abc", "String_LTrim"
    Test_Equal mod_string.String_RTrim("abc   "), "abc", "String_RTrim"
    Test_Equal canon(mod_string.String_To_Array("ab")), "a,b", "String_To_Array"
    Test_Equal mod_string.String_Join(Array("a", "b", "c"), "-"), "a-b-c", "String_Join"
    Test_True (Len(mod_string.String_Random(12)) = 12), "String_Random: 长度"
    Test_True mod_regex.Regex_Test("^[A-Za-z0-9]+$", mod_string.String_Random(20)), "String_Random: 字符集"
End Sub

Private Sub suite_regex()
    log_line "--- mod_regex ---"
    Test_True mod_regex.Regex_Test(".hello.", "--helloo--"), "Regex_Test: 匹配"
    Test_False mod_regex.Regex_Test("^\d+$", "ab12"), "Regex_Test: 不匹配"
    Test_Equal mod_regex.Regex_Find("\d\d年", "kkk2023年"), "23年", "Regex_Find: 首个匹配"
    Test_Equal mod_regex.Regex_Replace("\d\d年", "kkk2023年", "24年"), "kkk2024年", "Regex_Replace"
End Sub

Private Sub suite_date()
    log_line "--- mod_date ---"
    Test_True mod_regex.Regex_Test("^\d{8}$", mod_date.Date_Stamp("date")), "Date_Stamp(date): YYYYMMDD"
    Test_True mod_regex.Regex_Test("^\d{6}$", mod_date.Date_Stamp("time")), "Date_Stamp(time): hhmmss"
    Test_Equal (Len(mod_date.Date_Stamp("datetime"))), 14, "Date_Stamp(datetime): 长度"
    Test_Equal (Len(mod_date.Date_Stamp())), 16, "Date_Stamp()默认: 长度"
End Sub

Private Sub suite_file()
    log_line "--- mod_file ---"
    Dim tmp As String
    tmp = test_tmp_dir()
    Test_True mod_file.Folder_Exists(tmp), "Folder_Ensure / Folder_Exists"
    mod_file.File_Write tmp & "\a.txt", "hello"
    Test_True mod_file.File_Exists(tmp & "\a.txt"), "File_Write / File_Exists"
    Test_Equal mod_file.File_Name(tmp & "\a.txt"), "a.txt", "File_Name"
    Test_Equal mod_file.File_BaseName(tmp & "\a.txt"), "a", "File_BaseName"
    Test_Equal mod_file.File_ExtName(tmp & "\a.txt"), "txt", "File_ExtName"
    Test_True mod_file.File_Name_Valid("a_b-1.txt"), "File_Name_Valid: 合法"
    Test_False mod_file.File_Name_Valid("a:b.txt"), "File_Name_Valid: 非法字符"
    Test_False mod_file.File_Name_Valid(""), "File_Name_Valid: 空名"
    Dim files As Variant
    files = mod_file.Folder_ListFiles(tmp)
    Test_True mod_array.Array_Contains(files, tmp & "\a.txt"), "Folder_ListFiles: 包含 a.txt"
    ' 清理
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(tmp) Then fso.DeleteFolder tmp
    Set fso = Nothing
End Sub

Private Sub suite_dict()
    log_line "--- mod_dict ---"
    Dim d As Object
    Set d = mod_dict.Dict_Create()
    Test_Equal CStr(mod_dict.Dict_Count(d)), "0", "Dict_Count: 起始为 0"
    Test_False mod_dict.Dict_Exists(d, "a"), "Dict_Exists: 缺键"
    mod_dict.Dict_Set d, "a", 3
    mod_dict.Dict_Set d, "b", 5
    Test_Equal CStr(mod_dict.Dict_Get(d, "a")), "3", "Dict_Get: 命中"
    Test_Equal CStr(mod_dict.Dict_Get(d, "zz", 99)), "99", "Dict_Get: 缺键返回默认"
    Test_Equal CStr(mod_dict.Dict_Count(d)), "2", "Dict_Count: 记录数"
    Test_Equal canon(mod_dict.Dict_Keys(d)), "a,b", "Dict_Keys"
    Test_Equal canon(mod_dict.Dict_Values(d)), "3,5", "Dict_Values"
    Dim da As Variant
    da = mod_dict.Dict_To_Array(d)
    Test_Equal CStr(da(1, 1)) & "," & CStr(da(1, 2)), "a,3", "Dict_To_Array: 首行"
    mod_dict.Dict_Set d, "a", 9
    Test_Equal CStr(mod_dict.Dict_Get(d, "a")), "9", "Dict_Set: 覆盖已有键"
    mod_dict.Dict_Remove d, "b"
    Test_False mod_dict.Dict_Exists(d, "b"), "Dict_Remove: 移除后不存在"
    mod_dict.Dict_Clear d
    Test_Equal CStr(mod_dict.Dict_Count(d)), "0", "Dict_Clear"
End Sub

Private Sub suite_sort()
    log_line "--- mod_sort ---"
    Test_Equal canon(mod_sort.Array_Sort(Array(3, 1, 2))), "1,2,3", "Array_Sort: 升序"
    Test_Equal canon(mod_sort.Array_Sort(Array(3, 1, 2), False)), "3,2,1", "Array_Sort: 降序"
    Test_Equal canon(mod_sort.Array_Sort(Array("b", "a", "c"))), "a,b,c", "Array_Sort: 字符串"
    Test_Equal canon(mod_sort.Array_Reverse(Array(1, 2, 3))), "3,2,1", "Array_Reverse"
    Test_Equal canon(mod_sort.Array_Reverse(Array())), "", "Array_Reverse: 空数组"
    ' 稳定：按键 x mod 10 = (1,2,1,4) 升序，同键的 11(idx0) 应保持在 31(idx2) 之前
    Test_Equal canon(mod_sort.Array_Sort(Array(11, 2, 31, 4), True, "helper_key_mod10")), _
                "11,31,2,4", "Array_Sort: 稳定(按余数)"
End Sub

Private Sub suite_json()
    log_line "--- mod_json ---"
    Dim o As Variant
    Dim a As Variant
    Dim caught As Boolean

    ' 基础对象 / 字符串 / 转义（对象必须 Set 承接，Let 会触发字典 #450）
    Set o = mod_json.JSON_Parse("{""name"":""Alice"",""age"":30}")
    Test_Equal CStr(o("name")), "Alice", "JSON_Parse: 字符串值"
    Test_Equal CStr(o("age")), "30", "JSON_Parse: 整数值"

    ' 转义字符
    Set o = mod_json.JSON_Parse("{""nl"":""a\nb"",""q"":""say\""hi\""""}")
    Test_Equal o("nl"), "a" & vbLf & "b", "JSON_Parse: \n 转义"
    Test_Equal o("q"), "say""hi""", "JSON_Parse: 引号转义"

    ' Unicode 转义
    Set o = mod_json.JSON_Parse("{""c"":""\u4e2d\u6587""}")
    Test_Equal o("c"), "中文", "JSON_Parse: \\u 转义"

    ' 数组（数组是值，用 = 承接）
    a = mod_json.JSON_Parse("[10,20,-3]")
    Test_Equal CStr(a(0)), "10", "JSON_Parse: 数组元素[0]"
    Test_Equal CStr(a(2)), "-3", "JSON_Parse: 数组负数"

    ' 嵌套对象
    Set o = mod_json.JSON_Parse("{""user"":{""id"":7,""tags"":[""a"",""b""]}}")
    Test_Equal CStr(o("user")("id")), "7", "JSON_Parse: 嵌套对象取值"
    Test_Equal CStr(o("user")("tags")(1)), "b", "JSON_Parse: 嵌套数组取值"

    ' 布尔 / 空 / 浮点
    Set o = mod_json.JSON_Parse("{""t"":true,""f"":false,""n"":null,""pi"":3.5}")
    Test_True CBool(o("t")), "JSON_Parse: true"
    Test_False CBool(o("f")), "JSON_Parse: false"
    Test_True IsNull(o("n")), "JSON_Parse: null"
    Test_Equal CStr(o("pi")), "3.5", "JSON_Parse: 浮点数"

    ' 空对象 / 空数组
    Set o = mod_json.JSON_Parse("{}")
    Test_True (TypeName(o) = "Dictionary"), "JSON_Parse: 空对象"
    a = mod_json.JSON_Parse("[]")
    Test_True (UBound(a) < LBound(a)), "JSON_Parse: 空数组"

    ' 往返序列化
    Dim round As Variant
    Set round = mod_json.JSON_Parse("{""a"":1,""b"":[2,3]}")
    Test_Equal mod_json.JSON_Stringify(round), "{""a"":1,""b"":[2,3]}", "JSON 往返"
    Test_Equal mod_json.JSON_Stringify(Array(1, "x", True, Null)), "[1,""x"",true,null]", "JSON_Stringify 数组"
    Test_Equal mod_json.JSON_Stringify(Null), "null", "JSON_Stringify: null"

    ' 非法输入应抛错
    caught = False
    On Error Resume Next
    Dim dummy As Variant
    dummy = mod_json.JSON_Parse("{bad")
    If Err.Number <> 0 Then caught = True
    Err.Clear
    On Error GoTo 0
    Test_True caught, "JSON_Parse: 非法输入抛错"
End Sub

Private Sub suite_version()
    log_line "--- mod_version ---"
    Test_True (Len(mod_version.Version_String()) > 0), "Version_String: 非空"
    Test_True mod_regex.Regex_Test("^\d+\.\d+\.\d+$", mod_version.Version_String()), "Version_String: 符合 SemVer"
    Test_Equal mod_version.Version_String(), mod_version.LIB_VERSION, "Version_String == LIB_VERSION"
End Sub