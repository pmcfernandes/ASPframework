<!-- StoredProcedure.asp
   Stored Procedure wrapper for Classic ASP (VBScript)
   Usage:
     Dim sp: Set sp = New StoredProcedure
     sp.Init dbConn  ' pass an ADODB.Connection object OR a MSSQLConnection wrapper
     sp.Name = "usp_GetPerson"
     sp.AddParam "@Id", 3, 1, , 123 ' adInteger, adParamInput
     sp.AddParam "@Count", 3, 2 ' adInteger, adParamOutput
     Dim rs: Set rs = sp.Execute()
     Dim outCount: outCount = sp.GetOutput("@Count")
-->

<%
Class StoredProcedure
    Public Name
    Private connObj ' can be ADODB.Connection or MSSQLConnection
    Private cmd
    Private params

    Private Sub Class_Initialize()
        Name = ""
        Set params = CreateObject("Scripting.Dictionary")
        Set cmd = Nothing
    End Sub

    Public Sub Init(connection)
        ' Accept either an ADODB.Connection or MSSQLConnection wrapper
        Set connObj = connection
    End Sub

    Public Sub AddParam(name, dataType, direction, Optional size, Optional value)
        Dim p
        Set p = CreateObject("Scripting.Dictionary")
        p.Add "name", name
        p.Add "type", dataType
        p.Add "direction", direction
        If IsMissing(size) Then size = 0
        p.Add "size", size
        If Not IsMissing(value) Then p.Add "value", value Else p.Add "value", Null
        params.Add name, p
    End Sub

    Private Function GetActiveConnection()
        On Error Resume Next
        Dim c
        If TypeName(connObj) = "ADODB.Connection" Then
            Set GetActiveConnection = connObj
            Exit Function
        End If
        If TypeName(connObj) = "MSSQLConnection" Or TypeName(connObj) = "Object" Then
            ' MSSQLConnection wrapper
            ' Try to extract connection by opening a new one using its ConnectionString
            Set c = CreateObject("ADODB.Connection")
            c.Open connObj.ConnectionString
            Set GetActiveConnection = c
            Exit Function
        End If
        Set GetActiveConnection = Nothing
    End Function

    Public Function Execute()
        On Error Resume Next
        Dim conn, rs, k, p, direction
        Set Execute = Nothing
        Set conn = GetActiveConnection()
        If conn Is Nothing Then Exit Function

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn
        cmd.CommandText = Name
        cmd.CommandType = 4 ' adCmdStoredProc

        For Each k In params.Keys
            Set p = params.Item(k)
            direction = p.Item("direction")
            If p.Item("size") > 0 Then
                cmd.Parameters.Append cmd.CreateParameter(p.Item("name"), p.Item("type"), direction, p.Item("size"), p.Item("value"))
            Else
                cmd.Parameters.Append cmd.CreateParameter(p.Item("name"), p.Item("type"), direction, , p.Item("value"))
            End If
        Next

        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            ' error
            Err.Clear
            Set Execute = Nothing
            ' close connection if we opened it
            If TypeName(connObj) <> "ADODB.Connection" Then
                On Error Resume Next
                conn.Close
            End If
            Exit Function
        End If

        ' retrieve output params back into params dictionary
        For Each k In params.Keys
            Set p = params.Item(k)
            If p.Item("direction") = 2 Or p.Item("direction") = 4 Then ' adParamOutput or adParamReturnValue
                p.Add "out", cmd.Parameters(p.Item("name")).Value
            End If
        Next

        Set Execute = rs
    End Function

    Public Function GetOutput(name)
        If params.Exists(name) Then
            Dim p
            Set p = params.Item(name)
            If p.Exists("out") Then GetOutput = p.Item("out") Else GetOutput = Null
        Else
            GetOutput = Null
        End If
    End Function

    Private Sub Class_Terminate()
        On Error Resume Next
        If Not cmd Is Nothing Then Set cmd = Nothing
        Set connObj = Nothing
        Set params = Nothing
    End Sub
End Class
%>
