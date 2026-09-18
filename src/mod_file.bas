Attribute VB_Name = "mod_file"
'=====================================================================
' mod_file - 文件 / 文件夹 通用工具（通用库 · v1.0）
'=====================================================================
' 目录 / Catalog
'   File_Exists(path)                   As Boolean   文件是否存在
'   Folder_Exists(path)                 As Boolean   文件夹是否存在
'   Folder_Ensure(path)                 As String    确保文件夹存在(自动创建)，返回 path
'   File_Copy(from, to[, overwrite])              复制文件
'   File_Write(path, text[, append])              写入/追加文本(注意:默认 unicode=True 为 UTF-16,非 UTF-8)
'   File_Name(path)                     As String    文件名(含扩展名)
'   File_BaseName(path)                 As String    文件名(不含扩展名)
'   File_ExtName(path)                  As String    扩展名(不含点)
'   File_Name_Valid(name)               As Boolean   文件名是否合法
'   Folder_ListFiles(path[, recursive]) As Variant   列出文件全路径数组
'=====================================================================
Option Explicit

Private Function Get_FSO() As Object
    Static fso As Object
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    Set Get_FSO = fso
End Function

Public Function File_Exists(ByVal file_path As String) As Boolean
    File_Exists = Get_FSO().FileExists(file_path)
End Function

Public Function Folder_Exists(ByVal folder_path As String) As Boolean
    Folder_Exists = Get_FSO().FolderExists(folder_path)
End Function

Public Function Folder_Ensure(ByVal folder_path As String) As String
    Dim fso As Object
    Set fso = Get_FSO()
    If Not fso.FolderExists(folder_path) Then fso.CreateFolder folder_path
    Folder_Ensure = folder_path
End Function

Public Sub File_Copy(ByVal from_path As String, ByVal to_path As String, Optional ByVal overwrite As Boolean = True)
    Get_FSO().CopyFile from_path, to_path, overwrite
End Sub

' unicode=True 写 UTF-16(Unicode);False 写系统默认(ANSI)。
' 注意：FSO CreateTextFile 第3参 unicode=True 表示 UTF-16，并非 UTF-8；
' 配合 run_tests/vba_excel 读取报告时用 utf-16 解码，二者一致。
Public Sub File_Write(ByVal file_path As String, ByVal text As String, _
                      Optional ByVal append As Boolean = False, _
                      Optional ByVal unicode As Boolean = True)
    Dim fso As Object
    Dim tf As Object
    Set fso = Get_FSO()
    If fso.FileExists(file_path) And append Then
        Set tf = fso.OpenTextFile(file_path, 8, False)
    Else
        Set tf = fso.CreateTextFile(file_path, True, unicode)
    End If
    tf.Write text
    tf.Close
End Sub

Public Function File_Name(ByVal file_path As String) As String
    File_Name = Get_FSO().GetFileName(file_path)
End Function

Public Function File_BaseName(ByVal file_path As String) As String
    Dim n As String
    Dim dot As Long
    n = Get_FSO().GetFileName(file_path)
    dot = InStrRev(n, ".")
    If dot = 0 Then
        File_BaseName = n
    Else
        File_BaseName = Left$(n, dot - 1)
    End If
End Function

Public Function File_ExtName(ByVal file_path As String) As String
    Dim n As String
    Dim dot As Long
    n = Get_FSO().GetFileName(file_path)
    dot = InStrRev(n, ".")
    If dot = 0 Then
        File_ExtName = vbNullString
    Else
        File_ExtName = Mid$(n, dot + 1)
    End If
End Function

Public Function File_Name_Valid(ByVal name As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(name) = 0 Then
        File_Name_Valid = False
        Exit Function
    End If
    For i = 1 To Len(name)
        ch = Mid$(name, i, 1)
        If InStr(1, "\/:*?" & Chr$(34) & "<>|", ch, vbBinaryCompare) > 0 Then
            File_Name_Valid = False
            Exit Function
        End If
    Next i
    File_Name_Valid = True
End Function

Public Function Folder_ListFiles(ByVal folder_path As String, Optional ByVal recursive As Boolean = False) As Variant
    Dim col As Collection
    Dim fso As Object
    Set fso = Get_FSO()
    If Not fso.FolderExists(folder_path) Then
        Folder_ListFiles = Array()
        Exit Function
    End If
    Set col = New Collection
    Walk_Folder folder_path, recursive, col
    If col.Count = 0 Then
        Folder_ListFiles = Array()
    Else
        Folder_ListFiles = mod_array.Collection_To_Array(col)
    End If
End Function

Private Sub Walk_Folder(ByVal folder_path As String, ByVal recursive As Boolean, ByRef out_col As Collection)
    Dim fso As Object
    Dim f As Object
    Dim sf As Object
    Set fso = Get_FSO()
    For Each f In fso.GetFolder(folder_path).Files
        out_col.Add f.Path
    Next f
    If recursive Then
        For Each sf In fso.GetFolder(folder_path).SubFolders
            Walk_Folder sf.Path, True, out_col
        Next sf
    End If
End Sub