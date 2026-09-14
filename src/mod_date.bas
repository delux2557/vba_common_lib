Attribute VB_Name = "mod_date"
'=====================================================================
' mod_date - 日期时间 通用工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   Date_String()              As String   YYYYMMDD
'   Time_String()              As String   hhmmss
'   DateTime_String()          As String   YYYY_MMDD_hhmm
'   Full_DateTime_String()     As String   YYYY_MMDD_hhmmss
'=====================================================================
Option Explicit

Public Function Date_String() As String
    Date_String = Format$(Now, "YYYYMMDD")
End Function

Public Function Time_String() As String
    Time_String = Format$(Now, "hhmmss")
End Function

Public Function DateTime_String() As String
    DateTime_String = Format$(Now, "YYYY_MMDD_hhmm")
End Function

Public Function Full_DateTime_String() As String
    Full_DateTime_String = Format$(Now, "YYYY_MMDD_hhmmss")
End Function