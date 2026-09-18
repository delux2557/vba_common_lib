Attribute VB_Name = "mod_date"
'=====================================================================
' mod_date - 日期时间 通用工具（通用库 · v1.0）
'=====================================================================
' 职责：统一"取当前时间戳"。默认 stamp 为文件名安全格式；
'       其余 kind 供文件名/目录名/日志等不同粒度复用，勿再用散落的 Format$(Now...)。
' 目录 / Catalog
'   Date_Stamp([kind])      As String  当前时间戳；kind 决定输出格式
'                                       kind: date|time|datetime|stamp(默认)
'=====================================================================
Option Explicit

' 统一取"当前时间戳"并按 kind 格式化；默认 stamp 为文件名安全格式。
Public Function Date_Stamp(Optional ByVal kind As String = "stamp") As String
    Dim k As String
    k = LCase$(kind)
    Select Case k
        Case "date":     Date_Stamp = Format$(Now, "YYYYMMDD")              ' 20251031
        Case "time":     Date_Stamp = Format$(Now, "hhmmss")                ' 143005
        Case "datetime": Date_Stamp = Format$(Now, "YYYY_MMDD_hhmm")        ' 2025_1031_1430
        Case Else:       Date_Stamp = Format$(Now, "YYYY_MMDD_hhmmss")      ' 2025_1031_143005
    End Select
End Function