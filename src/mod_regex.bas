Attribute VB_Name = "mod_regex"
'=====================================================================
' mod_regex - 正则表达式 通用工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   Regex_Test(pattern, source)      As Boolean   是否匹配
'   Regex_Find(pattern, source)      As String    返回第一个匹配值（不含匹配则空串）
'   Regex_Replace(pattern, source, repl) As String 全局替换
' 说明：每个调用各自新建 RegExp（无共享状态 / 线程安全）；用 CreateObject
'       晚绑定，无需勾选 "Microsoft VBScript Regular Expressions" 引用。
'=====================================================================
Option Explicit

Private Function New_RegExp(ByVal pattern As String, ByVal ignore_case As Boolean) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = ignore_case
    re.MultiLine = True
    re.Pattern = pattern
    Set New_RegExp = re
End Function

Public Function Regex_Test(ByVal pattern As String, ByVal source As String) As Boolean
    Regex_Test = New_RegExp(pattern, True).Test(source)
End Function

Public Function Regex_Find(ByVal pattern As String, ByVal source As String) As String
    Dim re As Object
    Dim m As Variant
    Set re = New_RegExp(pattern, False)
    For Each m In re.Execute(source)
        Regex_Find = m.Value
        Exit Function
    Next m
    Regex_Find = vbNullString
End Function

Public Function Regex_Replace(ByVal pattern As String, ByVal source As String, ByVal repl As String) As String
    Regex_Replace = New_RegExp(pattern, True).Replace(source, repl)
End Function