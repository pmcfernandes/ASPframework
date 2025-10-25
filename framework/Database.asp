<!--#include virtual="/" -->
<!--
  Database.asp
  Classic ASP (VBScript) wrapper for MS SQL Server

  Usage:
    <!--#include file="Database.asp" -->
    Dim db: Set db = New MSSQLConnection
    db.ConnectionString = "Provider=SQLOLEDB;Data Source=SERVERNAME;Initial Catalog=DBNAME;User ID=sa;Password=yourpassword;"
    db.Connect
    Dim rs: Set rs = db.ExecuteQuery("SELECT TOP 1 * FROM SomeTable")
    If Not rs.EOF Then Response.Write rs(0)
    rs.Close
    Set rs = Nothing
    db.Close
    Set db = Nothing

  The class supports transaction methods and basic error handling.
-->

<%
'---------------------------------------------
' MSSQLConnection class
'---------------------------------------------
Class MSSQLConnection
    Public ConnectionString
    Private conn
    Private lastError

    Private Sub Class_Initialize()
        ConnectionString = ""
        Set conn = Server.CreateObject("ADODB.Connection")
        lastError = ""
    End Sub

    Public Function Connect()
        On Error Resume Next
        If conn.State = 1 Then
            Connect = True
            Exit Function
        End If

        conn.Open ConnectionString
        If Err.Number <> 0 Then
            lastError = "Connect error: " & Err.Number & " - " & Err.Description
            Connect = False
            Err.Clear
            Exit Function
        End If

        Connect = True
    End Function

    Public Sub Close()
        On Error Resume Next
        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
            Set conn = Nothing
        End If
    End Sub

    Public Function ExecuteQuery(sql)
        On Error Resume Next
        Dim rs
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.CursorLocation = 3 ' adUseClient
        rs.Open sql, conn, 3, 3 ' adOpenStatic, adLockReadOnly

        If Err.Number <> 0 Then
            lastError = "ExecuteQuery error: " & Err.Number & " - " & Err.Description
            Err.Clear
            Set ExecuteQuery = Nothing
            Exit Function
        End If

        Set ExecuteQuery = rs
    End Function

    Public Function ExecuteNonQuery(sql)
        On Error Resume Next
        Dim affected
        conn.Execute sql, affected
        If Err.Number <> 0 Then
            lastError = "ExecuteNonQuery error: " & Err.Number & " - " & Err.Description
            ExecuteNonQuery = -1
            Err.Clear
            Exit Function
        End If

        ExecuteNonQuery = affected
    End Function

    Public Function GetScalar(sql)
        On Error Resume Next
        Dim rs, val
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, conn, 3, 3
        If Err.Number <> 0 Then
            lastError = "GetScalar error: " & Err.Number & " - " & Err.Description
            GetScalar = Null
            Err.Clear
            Exit Function
        End If

        If Not rs.EOF Then
            val = rs(0)
        Else
            val = Null
        End If

        rs.Close
        Set rs = Nothing
        GetScalar = val
    End Function

    Public Sub BeginTrans()
        On Error Resume Next
        conn.BeginTrans
        If Err.Number <> 0 Then
            lastError = "BeginTrans error: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
    End Sub

    Public Sub CommitTrans()
        On Error Resume Next
        conn.CommitTrans
        If Err.Number <> 0 Then
            lastError = "CommitTrans error: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
    End Sub

    Public Sub RollbackTrans()
        On Error Resume Next
        conn.RollbackTrans
        If Err.Number <> 0 Then
            lastError = "RollbackTrans error: " & Err.Number & " - " & Err.Description
            Err.Clear
        End If
    End Sub

    Public Function GetLastError()
        GetLastError = lastError
    End Function

    Private Sub Class_Terminate()
        On Error Resume Next
        If Not conn Is Nothing Then
            If conn.State = 1 Then conn.Close
            Set conn = Nothing
        End If
    End Sub
End Class
%>
