Attribute VB_Name = "mod_string"
'=====================================================================
' mod_string - 字符串 通用工具（通用库 · v1.0）
'=====================================================================
' 职责：字符串与数组互转、连接、裁剪、随机生成、判前后缀、填充。
' 约定：去空白默认按"空格/制表/回车/换行"，也可用 trim_chars 指定任意字符集。
'
' 目录 / Catalog
'   String_To_Array(s)                 As Variant   字符串 -> 单字符数组
'   String_Join(arr[, delim])          As String    数组 -> 字符串（默认无分隔）
'   String_Trim(s[, chars])            As String    去首尾（默认空白;可指定字符集）
'   String_LTrim(s[, chars])           As String    去开头
'   String_RTrim(s[, chars])           As String    去结尾
'   String_Random(len[, char_set])     As String    随机字符串
'   String_StartsWith(s, prefix)       As Boolean   是否以 prefix 开头
'   String_EndsWith(s, suffix)         As Boolean   是否以 suffix 结尾
'   String_LeftPad(s, len[, pad_char]) As String    左填充到指定长度
'   String_Format(tpl, args...)       As String    模板格式化：{0}/{1} 占位符替换，{{ }} 转义
'=====================================================================
Option Explicit

' 按单个字符把字符串拆成 0 基数组；空串返回 Array()。
Public Function String_To_Array(ByVal s As String) As Variant
    Dim i As Long
    Dim arr() As Variant
    If Len(s) = 0 Then
        String_To_Array = Array()
        Exit Function
    End If
    ReDim arr(0 To Len(s) - 1)
    For i = 0 To Len(s) - 1
        arr(i) = Mid$(s, i + 1, 1)
    Next i
    String_To_Array = arr
End Function

' 把数组各元素用 delim 连接成字符串（默认无分隔）。元素统一 CStr，空数组合 ""。
Public Function String_Join(ByRef arr As Variant, Optional ByVal delim As String = vbNullString) As String
    Dim elem As Variant
    Dim first As Boolean
    first = True
    For Each elem In arr
        If Not first Then String_Join = String_Join & delim
        String_Join = String_Join & CStr(elem)
        first = False
    Next elem
End Function

' 去掉首尾 trim_chars 中的字符（缺省即空白集），等价于 LTrim+RTrim。
Public Function String_Trim(ByVal s As String, Optional ByVal trim_chars As String = vbNullString) As String
    String_Trim = String_LTrim(String_RTrim(s, trim_chars), trim_chars)
End Function

' 去掉开头连续的 trim_chars 字符；全被裁掉时返回空串。
Public Function String_LTrim(ByVal s As String, Optional ByVal trim_chars As String = vbNullString) As String
    Dim cset As String
    Dim p As Long
    cset = trim_set(trim_chars)
    For p = 1 To Len(s)
        If InStr(1, cset, Mid$(s, p, 1), vbBinaryCompare) = 0 Then
            String_LTrim = Mid$(s, p)
            Exit Function
        End If
    Next p
    String_LTrim = vbNullString
End Function

' 去掉结尾连续的 trim_chars 字符；全被裁掉时返回空串。
Public Function String_RTrim(ByVal s As String, Optional ByVal trim_chars As String = vbNullString) As String
    Dim cset As String
    Dim p As Long
    cset = trim_set(trim_chars)
    For p = Len(s) To 1 Step -1
        If InStr(1, cset, Mid$(s, p, 1), vbBinaryCompare) = 0 Then
            String_RTrim = Left$(s, p)
            Exit Function
        End If
    Next p
    String_RTrim = vbNullString
End Function

' 生成长度为 str_len 的随机字符串；默认字符集为大小写字母+数字，可传 char_set 自定义。
Public Function String_Random(ByVal str_len As Long, Optional ByVal char_set As String = vbNullString) As String
    Dim sb As String
    Dim i As Long
    Dim pool As String
    If Len(char_set) = 0 Then
        pool = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    Else
        pool = char_set
    End If
    Randomize
    For i = 1 To str_len
        sb = sb & Mid$(pool, Int(Len(pool) * Rnd) + 1, 1)
    Next i
    String_Random = sb
End Function

' 是否以 prefix 开头（区分大小写；prefix 为空视为 True）。
Public Function String_StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    String_StartsWith = (Left$(s, Len(prefix)) = prefix)
End Function

' 是否以 suffix 结尾（区分大小写；suffix 为空视为 True）。
Public Function String_EndsWith(ByVal s As String, ByVal suffix As String) As Boolean
    String_EndsWith = (Right$(s, Len(suffix)) = suffix)
End Function

' 左填充到总长 total_len：不足用 pad_char（取首字符）补齐；已达目标长度则原样返回。
Public Function String_LeftPad(ByVal s As String, ByVal total_len As Long, _
                               Optional ByVal pad_char As String = " ") As String
    If Len(s) >= total_len Then
        String_LeftPad = s
        Exit Function
    End If
    If Len(pad_char) < 1 Then pad_char = " "
    String_LeftPad = String$(total_len - Len(s), Left$(pad_char, 1)) & s
End Function

' 模板格式化：把 {0}/{1}/... 占位符替换为 args 对应参数；{{ 与 }} 转义为字面花括号。
' 参数越界 / 占位符语法错误 / 多余右花括号均抛错 #45000（调用方错误尽早暴露）。
Public Function String_Format(ByVal template As String, ParamArray args()) As String
    Dim out As String
    Dim i As Long
    Dim p As Long
    Dim c As String
    Dim num_buf As String
    Dim idx As Long
    Dim n_args As Long
    ' ParamArray 为空时 UBound 抛错；探测得到 -1 表示无参数
    On Error Resume Next
    n_args = UBound(args)
    If Err.Number <> 0 Then n_args = -1
    Err.Clear
    On Error GoTo 0

    out = ""
    i = 1
    Do While i <= Len(template)
        c = Mid$(template, i, 1)
        If c = "{" Then
            If Mid$(template, i + 1, 1) = "{" Then
                out = out & "{"
                i = i + 2
            Else
                num_buf = ""
                p = i + 1
                Do While p <= Len(template) And Mid$(template, p, 1) >= "0" And Mid$(template, p, 1) <= "9"
                    num_buf = num_buf & Mid$(template, p, 1)
                    p = p + 1
                Loop
                If Len(num_buf) = 0 Or Mid$(template, p, 1) <> "}" Then
                    Err.Raise 45000, "mod_string", "String_Format: 占位符语法错误 @ " & i
                End If
                On Error Resume Next
                idx = CLng(num_buf)
                If Err.Number <> 0 Then
                    Err.Clear
                    Err.Raise 45000, "mod_string", "String_Format: 占位符序号过大 @ " & i
                End If
                On Error GoTo 0
                If idx > n_args Then Err.Raise 45000, "mod_string", "String_Format: 参数越界 " & idx
                out = out & CStr(args(idx))
                i = p + 1
            End If
        ElseIf c = "}" Then
            If Mid$(template, i + 1, 1) = "}" Then
                out = out & "}"
                i = i + 2
            Else
                Err.Raise 45000, "mod_string", "String_Format: 多余的 '}' @ " & i
            End If
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    String_Format = out
End Function

' 默认去空格的字符集：空格 / 制表 / 回车 / 换行
Private Function trim_set(ByVal specified As String) As String
    If Len(specified) > 0 Then
        trim_set = specified
    Else
        trim_set = " " & vbTab & vbCr & vbLf
    End If
End Function