<!-- EnvHelper.asp
   Simple .env parser and accessor for Classic ASP
   Usage:
     LoadEnv Server.MapPath("/.env")
     Response.Write GetEnv("DB_SERVER")
-->

<%
Dim __env_vars
Set __env_vars = CreateObject("Scripting.Dictionary")

Sub LoadEnv(path)
    On Error Resume Next
    Dim fso, file, ts, line, kv, idx, key, val
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Sub
    Set file = fso.GetFile(path)
    Set ts = file.OpenAsTextStream(1, -2)
    Do While Not ts.AtEndOfStream
        line = Trim(ts.ReadLine)
        If line <> "" And Left(line,1) <> "#" Then
            idx = InStr(line, "=")
            If idx > 0 Then
                key = Trim(Left(line, idx-1))
                val = Trim(Mid(line, idx+1))
                ' remove optional surrounding quotes
                If Left(val,1) = """" And Right(val,1) = """" Then val = Mid(val,2, Len(val)-2)
                If __env_vars.Exists(key) Then __env_vars.Remove key
                __env_vars.Add key, val
            End If
        End If
    Loop
    ts.Close
    Set ts = Nothing
    Set file = Nothing
    Set fso = Nothing
End Sub

Function GetEnv(key, Optional defaultValue)
    If __env_vars.Exists(key) Then GetEnv = __env_vars.Item(key) Else GetEnv = defaultValue
End Function

Function AllEnv()
    Set AllEnv = __env_vars
End Function

Function ReplaceVars(s)
    ' Replace ${VAR} occurrences with environment values
    Dim re, matches, m, name, i
    Set re = New RegExp
    re.Pattern = "\$\{([A-Za-z0-9_]+)\}"
    re.Global = True
    Set matches = re.Execute(s)
    For i = matches.Count - 1 To 0 Step -1
        Set m = matches(i)
        name = m.SubMatches(0)
        s = Left(s, m.FirstIndex) & CStr(GetEnv(name, "")) & Mid(s, m.FirstIndex + m.Length + 1)
    Next
    ReplaceVars = s
End Function

' auto-load .env in the application root if present
On Error Resume Next
Dim defaultEnvPath
defaultEnvPath = Server.MapPath("/.env")
If CreateObject("Scripting.FileSystemObject").FileExists(defaultEnvPath) Then
    LoadEnv defaultEnvPath
End If
%>
