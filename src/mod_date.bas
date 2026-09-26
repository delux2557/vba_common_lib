Attribute VB_Name = "mod_date"
'=====================================================================
' mod_date - 日期时间 通用工具（通用库 · v1.1）
'=====================================================================
' 职责：统一"取当前时间戳"。默认 stamp 为文件名安全格式；
'       其余 kind 供文件名/目录名/日志等不同粒度复用，勿再用散落的 Format$(Now...)。
' 目录 / Catalog
'   Date_Stamp([kind])          As String  当前时间戳；kind: date|time|datetime|stamp(默认)
'   Date_StampUtc([kind])       As String  UTC 时间戳（同款 kind 格式）
'   Date_UtcNow()               As Date    当前 UTC 时间（GetSystemTime，晚绑定无 COM 依赖）
'   Date_ToUtc(dt_local)        As Date    本地时间 -> UTC（按当前时区偏移）
'   Date_FromUtc(utc_dt)        As Date    UTC -> 本地时间
'   Date_Format(dt[, pattern])  As String  按模式格式化（默认 yyyy-mm-dd hh:nn:ss）
' 说明：
'   - 时区偏移取"当前时刻"的本地-UTC 差；冬夏令时切换前后各 1 小时内，
'     对跨切换点时刻的换算可能有 1 小时误差（工具场景可接受）。
'   - Date_Stamp / Date_StampUtc 底层共用 stamp_of，保证格式一致。
'=====================================================================
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

#If VBA7 Then
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#Else
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
#End If

' 统一取"当前时间戳"并按 kind 格式化；默认 stamp 为文件名安全格式。
Public Function Date_Stamp(Optional ByVal kind As String = "stamp") As String
    Date_Stamp = stamp_of(Now, kind)
End Function

' 统一取"当前 UTC 时间戳"，格式与 Date_Stamp 一致。
Public Function Date_StampUtc(Optional ByVal kind As String = "stamp") As String
    Date_StampUtc = stamp_of(Date_UtcNow(), kind)
End Function

' 当前 UTC 时间（Windows API GetSystemTime，无需 COM 对象）。
Public Function Date_UtcNow() As Date
    Dim st As SYSTEMTIME
    GetSystemTime st
    Date_UtcNow = DateSerial(st.wYear, st.wMonth, st.wDay) _
                + TimeSerial(st.wHour, st.wMinute, st.wSecond) _
                + st.wMilliseconds / 86400000#
End Function

' 本地时间 -> UTC：减去当前时刻的本地-UTC 时差。
' 注意：参数名不能用 local（VBA 保留字，会导致该过程无法编译）。
Public Function Date_ToUtc(ByVal dt_local As Date) As Date
    Date_ToUtc = dt_local - tz_offset()
End Function

' UTC -> 本地时间：加上当前时刻的本地-UTC 时差。
Public Function Date_FromUtc(ByVal utc_dt As Date) As Date
    Date_FromUtc = utc_dt + tz_offset()
End Function

' 按自定义模式格式化时间；默认 "yyyy-mm-dd hh:nn:ss"（统一默认格式的 Format$ 封装）。
Public Function Date_Format(ByVal dt As Date, _
                            Optional ByVal pattern As String = "yyyy-mm-dd hh:nn:ss") As String
    Date_Format = Format$(dt, pattern)
End Function

'--- 内部实现 ---------------------------------------------------------

Private Function stamp_of(ByVal dt As Date, ByVal kind As String) As String
    Dim k As String
    k = LCase$(kind)
    Select Case k
        Case "date":     stamp_of = Format$(dt, "YYYYMMDD")              ' 20251031
        Case "time":     stamp_of = Format$(dt, "hhmmss")                ' 143005
        Case "datetime": stamp_of = Format$(dt, "YYYY_MMDD_hhmm")        ' 2025_1031_1430
        Case Else:       stamp_of = Format$(dt, "YYYY_MMDD_hhmmss")      ' 2025_1031_143005
    End Select
End Function

' 当前时刻的本地-UTC 时差（UTC 早于本地为正，如中国 +8 小时）。
Private Function tz_offset() As Date
    tz_offset = Now - Date_UtcNow()
End Function
