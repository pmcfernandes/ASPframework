<!-- DebugHelper.asp
   Debug helper for Classic ASP (VBScript).
   Usage:
     <!--#include file="DebugHelper.asp" -->
     Dim dbg: Set dbg = New DebugHelper
     dbg.LogToScreen = True
     dbg.LogLevel = dbg.LEVEL_DEBUG
     dbg.Log("Starting upload...", dbg.LEVEL_INFO)
     dbg.DumpDict someDict, "myDict"
     dbg.DumpRS rs, "users"
     dbg.LogToFile Server.MapPath("/logs/app.log")
-->

<%
Class DebugHelper
    ' Log levels
    Public Const LEVEL_FATAL = 1
    Public Const LEVEL_ERROR = 2
    Public Const LEVEL_WARN  = 3
    Public Const LEVEL_INFO  = 4
    Public Const LEVEL_DEBUG = 5

    Public LogLevel ' numeric: only messages with level <= LogLevel are emitted
    Public LogToScreen ' boolean
    Private logFilePath ' path to file for append
    Private fso

    Private Sub Class_Initialize()
        LogLevel = LEVEL_INFO
        LogToScreen = False
        logFilePath = ""
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
    End Sub

    Public Sub LogToFile(path)
        If Len(Trim(path)) = 0 Then Exit Sub
        logFilePath = path
        ' ensure directory exists
        Dim folder
        folder = fso.GetParentFolderName(path)
        If Not fso.FolderExists(folder) Then fso.CreateFolder(folder)
    End Sub

    Public Sub Log(msg, Optional level)
        If IsMissing(level) Then level = LEVEL_INFO
        If level > LogLevel Then Exit Sub
        Dim ts
        ts = Now()
        Dim line
        line = "[" & CStr(ts) & "] [" & LevelName(level) & "] " & CStr(msg)
        If LogToScreen Then
            Response.Write Server.HTMLEncode(line) & "<br/>"
        End If
        If Len(logFilePath) > 0 Then
            On Error Resume Next
            Dim stm
            Set stm = Server.CreateObject("ADODB.Stream")
            stm.Type = 2 ' adTypeText
            stm.Charset = "utf-8"
            stm.Open
            stm.WriteText line & vbCrLf
            stm.SaveToFile logFilePath, 8 ' adSaveCreateNotExist + append - using 8 to append
            stm.Close
            Set stm = Nothing
        End If
    End Sub

    Private Function LevelName(l)
        Select Case l
            Case LEVEL_FATAL
                LevelName = "FATAL"
            Case LEVEL_ERROR
                LevelName = "ERROR"
            Case LEVEL_WARN
                LevelName = "WARN"
            Case LEVEL_INFO
                LevelName = "INFO"
            Case LEVEL_DEBUG
                LevelName = "DEBUG"
            Case Else
                LevelName = "LVL" & CStr(l)
        End Select
    End Function

    ' Dump a Scripting.Dictionary to screen/file
    Public Sub DumpDict(d, Optional title)
        If IsObject(d) = False Then
            Log "DumpDict called with non-object", LEVEL_WARN
            Exit Sub
        End If
        If IsMissing(title) Then title = "dict"
        Dim k
        Log "--- DumpDict: " & title & " ---", LEVEL_DEBUG
        For Each k In d.Keys
            Log CStr(k) & " = " & CStr(d.Item(k)), LEVEL_DEBUG
        Next
        Log "--- end DumpDict ---", LEVEL_DEBUG
    End Sub

    ' Dump an array (1-based or 0-based)
    Public Sub DumpArray(arr, Optional title)
        If IsMissing(title) Then title = "arr"
        If IsArray(arr) = False Then
            Log "DumpArray called with non-array", LEVEL_WARN
            Exit Sub
        End If
        Dim i
        On Error Resume Next
        For i = LBound(arr) To UBound(arr)
            Log CStr(i) & ": " & CStr(arr(i)), LEVEL_DEBUG
        Next
    End Sub

    ' Dump an ADODB.Recordset: prints column headers and first N rows (default 50)
    Public Sub DumpRS(rs, Optional title, Optional maxRows)
        If IsMissing(maxRows) Then maxRows = 50
        If IsMissing(title) Then title = "recordset"
        If IsObject(rs) = False Then
            Log "DumpRS called with non-object", LEVEL_WARN
            Exit Sub
        End If
        If rs.EOF And rs.BOF Then
            Log "Recordset empty", LEVEL_DEBUG
            Exit Sub
        End If
        Dim fld
        Dim colNames
        colNames = ""
        For Each fld In rs.Fields
            If Len(colNames) = 0 Then colNames = fld.Name Else colNames = colNames & ", " & fld.Name
        Next
        Log "--- DumpRS: " & title & " cols: " & colNames & " ---", LEVEL_DEBUG
        Dim r
        r = 0
        rs.MoveFirst
        Do While Not rs.EOF And r < maxRows
            Dim out
            out = ""
            For Each fld In rs.Fields
                If Len(out) = 0 Then out = fld.Value Else out = out & " | " & fld.Value
            Next
            Log out, LEVEL_DEBUG
            rs.MoveNext
            r = r + 1
        Loop
        Log "--- end DumpRS ---", LEVEL_DEBUG
    End Sub

    ' Convenience: print a stack-like trace (limited on classic ASP)
    Public Sub Trace(msg)
        Log "TRACE: " & msg, LEVEL_DEBUG
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        Set fso = Nothing
    End Sub
End Class
%>
