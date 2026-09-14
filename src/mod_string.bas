Attribute VB_Name = "mod_string"
'=====================================================================
' mod_string - 字符串 通用工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   String_To_Array(s)                 As Variant   字符串 -> 单字符数组
'   String_Join(arr[, delim])          As String    数组 -> 字符串（默认无分隔）
'   String_Trim(s[, chars])            As String    去首尾（默认空白;可指定字符集）
'   String_LTrim(s[, chars])           As String    去开头
'   String_RTrim(s[, chars])           As String    去结尾
'   String_Random(len[, char_set])     As String    随机字符串
'=====================================================================
Option Explicit

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

Public Function String_Trim(ByVal s As String, Optional ByVal trim_chars As String = vbNullString) As String
    String_Trim = String_LTrim(String_RTrim(s, trim_chars), trim_chars)
End Function

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

' 默认去空格的字符集：空格 / 制表 / 回车 / 换行
Private Function trim_set(ByVal specified As String) As String
    If Len(specified) > 0 Then
        trim_set = specified
    Else
        trim_set = " " & vbTab & vbCr & vbLf
    End If
End Function