<!-- Model.asp
   BaseModel for Classic ASP MVC stack
   Usage:
     Class Person
         Inherits BaseModel ' (VBScript doesn't support inherits; copy pattern below)
     End Class
-->

<!--#include file="env_helper.asp" -->
<!--#include file="Database.asp" -->
<!--#include file="rs_wrapper.asp" -->

<%
Class BaseModel
    Public TableName
    Public PrimaryKey

    Private Sub Class_Initialize()
        TableName = ""
        PrimaryKey = "ID"
    End Sub

    Private Function GetDB()
        Dim db
        Set db = New MSSQLConnection
        Dim provider, server, name, user, pass, connStr, envConn
        envConn = GetEnv("DB_CONN_STRING", "")
        If envConn <> "" Then
            connStr = ReplaceVars(envConn)
        Else
            provider = GetEnv("DB_PROVIDER", "SQLOLEDB")
            server = GetEnv("DB_SERVER", "localhost")
            name = GetEnv("DB_NAME", "master")
            user = GetEnv("DB_USER", "")
            pass = GetEnv("DB_PASS", "")
            connStr = "Provider=" & provider & ";Data Source=" & server & ";Initial Catalog=" & name & ";"
            If Trim(user) <> "" Then
                connStr = connStr & "User ID=" & user & ";Password=" & pass & ";"
            Else
                connStr = connStr & "Integrated Security=SSPI;"
            End If
        End If
        db.ConnectionString = connStr
        If db.Connect() Then
            Set GetDB = db
        Else
            Set GetDB = Nothing
        End If
    End Function

    Public Function FindById(id)
        Dim db, rs, sql, rw
        Set FindById = Nothing
        If TableName = "" Then Exit Function
        Set db = GetDB()
        If db Is Nothing Then Exit Function
        sql = "SELECT * FROM " & TableName & " WHERE " & PrimaryKey & " = " & CStr(id)
        Set rs = db.ExecuteQuery(sql)
        If Not rs Is Nothing Then
            If Not rs.EOF Then
                Set rw = RecordsetToDict(rs)
                Set FindById = rw
            End If
            rs.Close
            Set rs = Nothing
        End If
        db.Close
        Set db = Nothing
    End Function

    Public Function FindAll(Optional whereClause)
        Dim db, rs, sql, arr, idx
        If TableName = "" Then Exit Function
        Set FindAll = Nothing
        Set db = GetDB()
        If db Is Nothing Then Exit Function
        sql = "SELECT * FROM " & TableName
        If Not IsMissing(whereClause) And Trim(whereClause) <> "" Then sql = sql & " WHERE " & whereClause
        Set rs = db.ExecuteQuery(sql)
        If rs Is Nothing Then
            db.Close
            Set db = Nothing
            Exit Function
        End If
        idx = 0
        Do While Not rs.EOF
            If IsNull(FindAll) Then ReDim arr(0)
            Set arr(idx) = RecordsetToDict(rs)
            idx = idx + 1
            rs.MoveNext
            ReDim Preserve arr(idx-1)
        Loop
        rs.Close
        Set rs = Nothing
        db.Close
        Set db = Nothing
        Set FindAll = arr
    End Function

    Public Function Insert(dict)
        ' dict = Scripting.Dictionary of field->value
        Dim db, sql, keys, vals, k, cmd, placeholders
        Set Insert = Nothing
        If TableName = "" Then Exit Function
        Set db = GetDB()
        If db Is Nothing Then Exit Function
        keys = ""
        vals = ""
        placeholders = ""
        For Each k In dict.Keys
            If keys <> "" Then keys = keys & ","
            keys = keys & k
            If vals <> "" Then vals = vals & ","
            vals = vals & "?"
            If placeholders <> "" Then placeholders = placeholders & ","
            placeholders = placeholders & "?"
        Next
        sql = "INSERT INTO " & TableName & " (" & keys & ") VALUES (" & vals & ")"
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(db)
        cmd.CommandText = sql
        cmd.CommandType = 1
        For Each k In dict.Keys
            cmd.Parameters.Append cmd.CreateParameter("p", 200, 1, 8000, dict.Item(k))
        Next
        On Error Resume Next
        cmd.Execute
        If Err.Number = 0 Then Insert = True Else Insert = False
        If Not cmd Is Nothing Then Set cmd = Nothing
        db.Close
        Set db = Nothing
    End Function

    Public Function Update(id, dict)
        Dim db, sql, setClause, k, cmd
        Set Update = Nothing
        If TableName = "" Then Exit Function
        Set db = GetDB()
        If db Is Nothing Then Exit Function
        setClause = ""
        For Each k In dict.Keys
            If setClause <> "" Then setClause = setClause & ","
            setClause = setClause & k & " = ?"
        Next
        sql = "UPDATE " & TableName & " SET " & setClause & " WHERE " & PrimaryKey & " = ?"
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(db)
        cmd.CommandText = sql
        cmd.CommandType = 1
        For Each k In dict.Keys
            cmd.Parameters.Append cmd.CreateParameter("p", 200, 1, 8000, dict.Item(k))
        Next
        cmd.Parameters.Append cmd.CreateParameter("idp", 3, 1, , CInt(id))
        On Error Resume Next
        cmd.Execute
        If Err.Number = 0 Then Update = True Else Update = False
        If Not cmd Is Nothing Then Set cmd = Nothing
        db.Close
        Set db = Nothing
    End Function

    Public Function Delete(id)
        Dim db, sql
        Set Delete = Nothing
        If TableName = "" Then Exit Function
        Set db = GetDB()
        If db Is Nothing Then Exit Function
        sql = "DELETE FROM " & TableName & " WHERE " & PrimaryKey & " = " & CStr(id)
        On Error Resume Next
        db.ExecuteNonQuery sql
        If Err.Number = 0 Then Delete = True Else Delete = False
        db.Close
        Set db = Nothing
    End Function

    Private Function RecordsetToDict(rs)
        Dim d, i
        Set d = CreateObject("Scripting.Dictionary")
        For i = 0 To rs.Fields.Count - 1
            d.Add rs.Fields(i).Name, rs.Fields(i).Value
        Next
        Set RecordsetToDict = d
    End Function

End Class
%>
