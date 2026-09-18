Attribute VB_Name = "mod_json"
'=====================================================================
' mod_json - JSON 解析 / 序列化 通用工具（通用库 · v1.1）
'=====================================================================
' 目录 / Catalog
'   JSON_Parse(json_text)    As Variant    解析 JSON 字符串为嵌套结构
'   JSON_Stringify(value)    As String     把结构化数据序列化为 JSON 字符串
'=====================================================================
' 数据模型映射：
'   JSON 对象  -> Scripting.Dictionary（晚绑定，键保持插入顺序）
'   JSON 数组  -> 0 基 Variant 数组（IsArray 判断）
'   字符串     -> String；数字 -> 整数 Long / 小数 Double
'   true/false -> Boolean；null -> VBA Null
' 说明：
'   - 纯手写递归下降解析器，无任何外部引用，晚绑定 CreateObject，移植无依赖。
'   - 数字输出用 Str$ 保证小数点恒为 '.'，规避本机区域设置差异。
'   - 若无法解析，抛出错误 #45000，Description 含出错位置/原因。
'=====================================================================
Option Explicit

'--- 对外 API --------------------------------------------------------
' 解析 JSON 字符串为嵌套结构（数据模型见头部）。非法输入抛错 #45000，Description 含出错位置。
' 注意：对象解析为 Dictionary，承接结果须用 Set（见下方#450 说明）。
Public Function JSON_Parse(ByVal json_text As String) As Variant
    Dim pos As Long
    pos = 1
    Dim result As Variant
    Call parse_value(json_text, pos, result)
    skip_ws json_text, pos
    If pos <= Len(json_text) Then Err.Raise 45000, , "JSON: 结尾存在多余字符 @ 位置 " & pos
    ' Dictionary 有带参默认属性 .Item，用 Let(=) 承接对象会触发 #450；
    ' 因此对象引用必须 Set 传回，调用方同样用 Set 承接。
    If IsObject(result) Then
        Set JSON_Parse = result
    Else
        JSON_Parse = result
    End If
End Function

' 把结构化数据序列化为 JSON：对象=Dictionary、数组=数组、数字/字符串/布尔按类型直出。
' 顶级 null/未知类型输出 "null"。
Public Function JSON_Stringify(ByVal value As Variant) As String
    JSON_Stringify = ser(value)
End Function

'--- 解析器 ----------------------------------------------------------
' 对象引用经 Set 写入 out，标量/数组经 Let 写入，避免命中字典默认属性的 #450。
Private Sub parse_value(ByRef s As String, ByRef pos As Long, ByRef out As Variant)
    skip_ws s, pos
    If pos > Len(s) Then Err.Raise 45000, , "JSON: 意外结束 @ 位置 " & pos
    Dim ch As String
    ch = Mid$(s, pos, 1)
    Select Case ch
        Case "{": Set out = parse_object(s, pos)
        Case "[": out = parse_array(s, pos)
        Case """": out = parse_string(s, pos)
        Case "t": expect_word s, pos, "true":  out = True
        Case "f": expect_word s, pos, "false": out = False
        Case "n": expect_word s, pos, "null":  out = Null
        Case Else: out = parse_number(s, pos)
    End Select
End Sub

Private Sub skip_ws(ByRef s As String, ByRef pos As Long)
    Dim ch As String
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub expect_word(ByRef s As String, ByRef pos As Long, ByVal word As String)
    If Mid$(s, pos, Len(word)) = word Then
        pos = pos + Len(word)
    Else
        Err.Raise 45000, , "JSON: 期望 '" & word & "' @ 位置 " & pos
    End If
End Sub

Private Function parse_object(ByRef s As String, ByRef pos As Long) As Variant
    pos = pos + 1   ' 跳过 '{'
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim key As String
    Dim val As Variant
    skip_ws s, pos
    If pos <= Len(s) And Mid$(s, pos, 1) = "}" Then
        pos = pos + 1
        Set parse_object = dict
        Exit Function
    End If
    Do
        skip_ws s, pos
        If pos > Len(s) Or Mid$(s, pos, 1) <> """" Then
            Err.Raise 45000, , "JSON: 对象键须为字符串 @ 位置 " & pos
        End If
        key = parse_string(s, pos)
        skip_ws s, pos
        If pos > Len(s) Or Mid$(s, pos, 1) <> ":" Then Err.Raise 45000, , "JSON: 缺 ':' @ 位置 " & pos
        pos = pos + 1
        val = vbNullString
        Call parse_value(s, pos, val)
        If Not dict.Exists(key) Then dict.Add key, val Else dict(key) = val
        skip_ws s, pos
        If pos <= Len(s) And Mid$(s, pos, 1) = "," Then
            pos = pos + 1
        ElseIf pos <= Len(s) And Mid$(s, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        Else
            Err.Raise 45000, , "JSON: 对象缺 ',' 或 '}' @ 位置 " & pos
        End If
    Loop
    Set parse_object = dict
End Function

Private Function parse_array(ByRef s As String, ByRef pos As Long) As Variant
    pos = pos + 1   ' 跳过 '['
    Dim arr() As Variant
    Dim cnt As Long
    cnt = 0
    ReDim arr(0 To 0)
    skip_ws s, pos
    If pos <= Len(s) And Mid$(s, pos, 1) = "]" Then
        pos = pos + 1
        parse_array = Array()
        Exit Function
    End If
    Do
        If cnt > UBound(arr) Then ReDim Preserve arr(0 To cnt)
        arr(cnt) = vbNullString
        Call parse_value(s, pos, arr(cnt))
        cnt = cnt + 1
        skip_ws s, pos
        If pos <= Len(s) And Mid$(s, pos, 1) = "," Then
            pos = pos + 1
        ElseIf pos <= Len(s) And Mid$(s, pos, 1) = "]" Then
            pos = pos + 1
            Exit Do
        Else
            Err.Raise 45000, , "JSON: 数组缺 ',' 或 ']' @ 位置 " & pos
        End If
    Loop
    ReDim Preserve arr(0 To cnt - 1)
    parse_array = arr
End Function

Private Function parse_string(ByRef s As String, ByRef pos As Long) As String
    pos = pos + 1   ' 跳过开头 '"'
    Dim out As String
    out = ""
    Dim ch As String
    Do While pos <= Len(s)
        ch = Mid$(s, pos, 1)
        If ch = """" Then
            pos = pos + 1
            parse_string = out
            Exit Function
        ElseIf ch = "\" Then
            pos = pos + 1
            out = out & parse_escape(s, pos)
        Else
            out = out & ch
            pos = pos + 1
        End If
    Loop
    Err.Raise 45000, , "JSON: 字符串未闭合 @ 位置 " & pos
End Function

Private Function parse_escape(ByRef s As String, ByRef pos As Long) As String
    If pos > Len(s) Then Err.Raise 45000, , "JSON: 转义序列缺失"
    Dim c As String
    Dim hex4 As String
    c = Mid$(s, pos, 1)
    pos = pos + 1
    Select Case c
        Case """": parse_escape = """"
        Case "\": parse_escape = "\"
        Case "/": parse_escape = "/"
        Case "b": parse_escape = Chr(8)
        Case "f": parse_escape = Chr(12)
        Case "n": parse_escape = vbLf
        Case "r": parse_escape = vbCr
        Case "t": parse_escape = vbTab
        Case "u":
            hex4 = Mid$(s, pos, 4)
            If Len(hex4) < 4 Then Err.Raise 45000, , "JSON: \u 需 4 位十六进制"
            pos = pos + 4
            parse_escape = ChrW(CLng("&H" & hex4))
        Case Else: Err.Raise 45000, , "JSON: 未知转义 \" & c
    End Select
End Function

Private Function parse_number(ByRef s As String, ByRef pos As Long) As Variant
    Dim i As Long
    Dim token As String
    i = pos
    If pos <= Len(s) And Mid$(s, pos, 1) = "-" Then pos = pos + 1
    Do While pos <= Len(s) And ((Mid$(s, pos, 1) >= "0" And Mid$(s, pos, 1) <= "9") Or Mid$(s, pos, 1) = ".")
        pos = pos + 1
    Loop
    If pos <= Len(s) And (Mid$(s, pos, 1) = "e" Or Mid$(s, pos, 1) = "E") Then
        pos = pos + 1
        If pos <= Len(s) And (Mid$(s, pos, 1) = "+" Or Mid$(s, pos, 1) = "-") Then pos = pos + 1
        Do While pos <= Len(s) And (Mid$(s, pos, 1) >= "0" And Mid$(s, pos, 1) <= "9")
            pos = pos + 1
        Loop
    End If
    token = Mid$(s, i, pos - i)
    If token = "" Or token = "-" Then Err.Raise 45000, , "JSON: 数字格式错误 @ 位置 " & i
    If InStr(1, token, ".", vbBinaryCompare) > 0 Or InStr(1, token, "E", vbBinaryCompare) > 0 Then
        parse_number = CDbl(token)
    Else
        parse_number = CLng(token)
    End If
End Function

'--- 序列化器 --------------------------------------------------------
Private Function ser(ByVal value As Variant) As Variant
    Dim vt As Integer
    If IsObject(value) Then
        If TypeName(value) = "Dictionary" Then
            ser = "{" & ser_dict(value) & "}"
        Else
            ser = "null"
        End If
    ElseIf IsArray(value) Then
        ser = "[" & ser_arr(value) & "]"
    ElseIf IsNull(value) Or IsEmpty(value) Then
        ser = "null"
    Else
        vt = VarType(value)
        Select Case vt
            Case vbBoolean:  If value Then ser = "true" Else ser = "false"
            Case vbString:   ser = quote_string(CStr(value))
            Case Else:       ser = ser_number(value)
        End Select
    End If
End Function

Private Function ser_dict(ByVal dict As Object) As String
    Dim k As Variant
    Dim items() As String
    Dim n As Long
    n = 0
    ReDim items(0 To 0)
    For Each k In dict.keys
        If n > UBound(items) Then ReDim Preserve items(0 To n)
        items(n) = quote_string(CStr(k)) & ":" & ser(dict(k))
        n = n + 1
    Next k
    If n = 0 Then
        ser_dict = ""
    Else
        ReDim Preserve items(0 To n - 1)
        ser_dict = Join(items, ",")
    End If
End Function

Private Function ser_arr(ByRef value As Variant) As String
    Dim i As Long
    Dim items() As String
    Dim n As Long
    n = 0
    ReDim items(0 To 0)
    For i = LBound(value) To UBound(value)
        If n > UBound(items) Then ReDim Preserve items(0 To n)
        items(n) = ser(value(i))
        n = n + 1
    Next i
    If n = 0 Then
        ser_arr = ""
    Else
        ReDim Preserve items(0 To n - 1)
        ser_arr = Join(items, ",")
    End If
End Function

Private Function quote_string(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, Chr(8), "\b")
    s = Replace(s, Chr(12), "\f")
    quote_string = """" & s & """"
End Function

' 数字序列化：整数直出无小数点，浮点用 Str$（小数点恒为 '.'，规避区域设置）。
Private Function ser_number(ByVal value As Variant) As String
    Dim vt As Integer
    vt = VarType(value)
    If vt = vbLong Or vt = vbInteger Or vt = vbByte Or vt = vbDecimal Or vt = vbSingle Or vt = vbDouble Then
        If value = Fix(value) And CDbl(Abs(value)) < 2147483647# Then
            ser_number = CStr(CLng(value))
        Else
            ser_number = LTrim$(Str$(CDbl(value)))
        End If
    Else
        ser_number = CStr(value)
    End If
End Function