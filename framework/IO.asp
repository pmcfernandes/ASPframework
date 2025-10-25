<!-- IO.asp
   File and Directory helper functions for Classic ASP
-->

<%
' Directory helpers
Function DirExists(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    DirExists = fso.FolderExists(path)
    Set fso = Nothing
End Function

Function CreateDir(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        On Error Resume Next
        fso.CreateFolder path
        If Err.Number <> 0 Then
            CreateDir = False
            Err.Clear
        Else
            CreateDir = True
        End If
    Else
        CreateDir = True
    End If
    Set fso = Nothing
End Function

Function RemoveDir(path, Optional force)
    If IsMissing(force) Then force = False
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then
        On Error Resume Next
        fso.DeleteFolder path, force
        If Err.Number <> 0 Then
            RemoveDir = False
            Err.Clear
        Else
            RemoveDir = True
        End If
    Else
        RemoveDir = False
    End If
    Set fso = Nothing
End Function

' File helpers
Function FileExists(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
    Set fso = Nothing
End Function

Function ReadFile(path)
    Dim fso, file, ts, content
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        ReadFile = ""
        Exit Function
    End If
    Set file = fso.GetFile(path)
    Set ts = file.OpenAsTextStream(1, -2)
    content = ts.ReadAll
    ts.Close
    Set ts = Nothing
    Set file = Nothing
    Set fso = Nothing
    ReadFile = content
End Function

Function WriteFile(path, content, Optional overwrite)
    If IsMissing(overwrite) Then overwrite = True
    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) And Not overwrite Then
        WriteFile = False
        Exit Function
    End If
    Set ts = fso.OpenTextFile(path, 2, True, -2) ' ForWriting, create if not exists
    ts.Write content
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    WriteFile = True
End Function

Function AppendFile(path, content)
    Dim fso, ts
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 8, True, -2) ' ForAppending
    ts.Write content
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    AppendFile = True
End Function

Function ListFiles(path, Optional pattern)
    Dim fso, folder, files, col, f, arr, i
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        ListFiles = Array()
        Exit Function
    End If
    Set folder = fso.GetFolder(path)
    Set col = folder.Files
    i = 0
    ReDim arr(-1)
    For Each f In col
        If IsMissing(pattern) Or InStr(LCase(f.Name), LCase(pattern)) > 0 Then
            ReDim Preserve arr(i)
            arr(i) = f.Path
            i = i + 1
        End If
    Next
    ListFiles = arr
    Set col = Nothing
    Set folder = Nothing
    Set fso = Nothing
End Function

Function CopyFile(src, dest, Optional overwrite)
    If IsMissing(overwrite) Then overwrite = True
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.CopyFile src, dest, overwrite
    If Err.Number <> 0 Then
        CopyFile = False
        Err.Clear
    Else
        CopyFile = True
    End If
    Set fso = Nothing
End Function

Function MoveFile(src, dest)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.MoveFile src, dest
    If Err.Number <> 0 Then
        MoveFile = False
        Err.Clear
    Else
        MoveFile = True
    End If
    Set fso = Nothing
End Function

Function DeleteFile(path)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then
        On Error Resume Next
        fso.DeleteFile path, True
        If Err.Number <> 0 Then
            DeleteFile = False
            Err.Clear
        Else
            DeleteFile = True
        End If
    Else
        DeleteFile = False
    End If
    Set fso = Nothing
End Function

Function GetFileSize(path)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        GetFileSize = -1
        Exit Function
    End If
    Set f = fso.GetFile(path)
    GetFileSize = f.Size
    Set f = Nothing
    Set fso = Nothing
End Function

Function GetFileModified(path)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then
        GetFileModified = ""
        Exit Function
    End If
    Set f = fso.GetFile(path)
    GetFileModified = f.DateLastModified
    Set f = Nothing
    Set fso = Nothing
End Function

%>
