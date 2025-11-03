<!-- Table.asp
   Table wrapper convenience class for Classic ASP
   Usage:
     Dim t: Set t = New Table
     t.Init "Users", "UserID", dbWrapper
     Dim rows: rows = t.List("", Array(), "CreatedAt DESC", 50, 0)
-->

<%
<!--#include file="EnvHelper.asp" -->
<!--#include file="Database.asp" -->
<!--#include file="RecordsetWrapper.asp" -->

Class Table
    Public Name
    Public PrimaryKey
    Public DBWrapper

    Private Sub Class_Initialize()
        Name = ""
        PrimaryKey = "ID"
        Set DBWrapper = Nothing
    End Sub

    Public Sub Init(tableName, Optional pk, Optional dbw)
        Name = tableName
        If Not IsMissing(pk) Then PrimaryKey = pk
        If Not IsMissing(dbw) Then Set DBWrapper = dbw
    End Sub

    ' Return a connected MSSQLConnection (wrapper) or open a temporary one
    Private Function ensureWrapper()
        If Not DBWrapper Is Nothing Then
            Set ensureWrapper = DBWrapper
            Exit Function
        End If
        Dim db: Set db = New MSSQLConnection
        Dim envConn
        envConn = GetEnv("DB_CONN_STRING", "")
        If envConn <> "" Then
            db.ConnectionString = ReplaceVars(envConn)
        Else
            Dim provider, server, name, user, pass, connStr
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
            db.ConnectionString = connStr
        End If
        If db.Connect() Then
            Set ensureWrapper = db
        Else
            Set ensureWrapper = Nothing
        End If
    End Function

    ' List rows. whereClause optional, paramsArr optional, orderBy optional, limit optional, offset optional
    Public Function List(Optional whereClause, Optional paramsArr, Optional orderBy, Optional limit, Optional offset)
        Dim wrapper, sql, cmd, rs, arr, idx
        Set List = Nothing
        If Name = "" Then Exit Function
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        sql = "SELECT * FROM " & Name
        If Not IsMissing(whereClause) And Trim(CStr(whereClause)) <> "" Then sql = sql & " WHERE " & whereClause
        If Not IsMissing(orderBy) And Trim(CStr(orderBy)) <> "" Then sql = sql & " ORDER BY " & orderBy
        If Not IsMissing(limit) And IsNumeric(limit) Then
            ' Simple TOP fallback for SQL Server if no offset requested
            If IsMissing(offset) Or Not IsNumeric(offset) Or CInt(offset) = 0 Then
                sql = "SELECT TOP " & CInt(limit) & " * FROM " & Name & (IIf(Trim(CStr(whereClause)) <> "", " WHERE " & whereClause, "")) & (IIf(Trim(CStr(orderBy)) <> "", " ORDER BY " & orderBy, ""))
                ' Reuse previous sql path
            Else
                ' For more complex offset support, user should provide their own SQL
            End If
        End If

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = sql
        cmd.CommandType = 1

        If Not IsMissing(paramsArr) Then
            On Error Resume Next
            Dim ub, i, pval, dtype, p
            ub = -1
            ub = UBound(paramsArr)
            If Err.Number = 0 Then
                For i = 0 To ub
                    pval = paramsArr(i)
                    If IsDate(pval) Then dtype = 7 ElseIf IsNumeric(pval) Then If InStr(CStr(pval), ".") > 0 Then dtype = 5 Else dtype = 3 Else dtype = 200
                    Set p = cmd.CreateParameter("p" & i, dtype, 1, , pval)
                    cmd.Parameters.Append p
                Next
            End If
            On Error GoTo 0
        End If

        On Error Resume Next
        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            List = Nothing
            Err.Clear
            Exit Function
        End If

        idx = 0
        Do While Not rs.EOF
            If IsNull(List) Then ReDim arr(0)
            Set arr(idx) = RecordsetToDict(rs)
            idx = idx + 1
            rs.MoveNext
            ReDim Preserve arr(idx-1)
        Loop

        rs.Close
        Set rs = Nothing
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
        Set List = arr
    End Function

    Public Function Count(Optional whereClause, Optional paramsArr)
        Dim wrapper, sql, cmd, rs, cnt
        Set Count = 0
        If Name = "" Then Exit Function
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        sql = "SELECT COUNT(1) AS c FROM " & Name
        If Not IsMissing(whereClause) And Trim(CStr(whereClause)) <> "" Then sql = sql & " WHERE " & whereClause
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = sql
        cmd.CommandType = 1
        If Not IsMissing(paramsArr) Then
            On Error Resume Next
            Dim ub, i, pval, dtype, p
            ub = -1
            ub = UBound(paramsArr)
            If Err.Number = 0 Then
                For i = 0 To ub
                    pval = paramsArr(i)
                    If IsDate(pval) Then dtype = 7 ElseIf IsNumeric(pval) Then If InStr(CStr(pval), ".") > 0 Then dtype = 5 Else dtype = 3 Else dtype = 200
                    Set p = cmd.CreateParameter("p" & i, dtype, 1, , pval)
                    cmd.Parameters.Append p
                Next
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            Count = 0
            Err.Clear
            Exit Function
        End If
        If Not rs.EOF Then cnt = rs("c") Else cnt = 0
        rs.Close
        Set rs = Nothing
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
        Count = cnt
    End Function

    Public Function FindById(id)
        Dim whereClause, rows
        whereClause = PrimaryKey & " = ?"
        rows = List(whereClause, Array(id))
        If IsNull(rows) Then
            Set FindById = Nothing
            Exit Function
        End If
        If UBound(rows) >= 0 Then
            Set FindById = rows(0)
        Else
            Set FindById = Nothing
        End If
    End Function

    ' Execute arbitrary parameterized SQL and return array of dict rows
    Public Function Query(baseSql, Optional paramsArr)
        Dim wrapper, cmd, rs, arr, idx
        Set Query = Nothing
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = baseSql
        cmd.CommandType = 1
        If Not IsMissing(paramsArr) Then
            On Error Resume Next
            Dim ub, i, pval, dtype, p
            ub = -1
            ub = UBound(paramsArr)
            If Err.Number = 0 Then
                For i = 0 To ub
                    pval = paramsArr(i)
                    If IsDate(pval) Then dtype = 7 ElseIf IsNumeric(pval) Then If InStr(CStr(pval), ".") > 0 Then dtype = 5 Else dtype = 3 Else dtype = 200
                    Set p = cmd.CreateParameter("p" & i, dtype, 1, , pval)
                    cmd.Parameters.Append p
                Next
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            Query = Nothing
            Err.Clear
            Exit Function
        End If
        idx = 0
        Do While Not rs.EOF
            If IsNull(Query) Then ReDim arr(0)
            Set arr(idx) = RecordsetToDict(rs)
            idx = idx + 1
            rs.MoveNext
            ReDim Preserve arr(idx-1)
        Loop
        rs.Close
        Set rs = Nothing
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
        Set Query = arr
    End Function

    ' PagedList: returns a dictionary with keys: items (array), total (integer), page (pageNumber), pageSize
    ' baseSql should NOT include ORDER BY, as OFFSET/FETCH requires ORDER BY. Provide orderBy separately.
    Public Function PagedList(baseSql, Optional paramsArr, Optional orderBy, Optional pageSize, Optional pageNumber)
        Dim result, items, total, offset, pagedSql, countSql, wrapper, cmd, rs, idx
        Set PagedList = Nothing
        If Trim(CStr(orderBy)) = "" Then
            ' OFFSET/FETCH requires ORDER BY; fallback to PrimaryKey
            orderBy = PrimaryKey & " ASC"
        End If
        If IsMissing(pageSize) Or Not IsNumeric(pageSize) Then pageSize = 20
        If IsMissing(pageNumber) Or Not IsNumeric(pageNumber) Then pageNumber = 1
        offset = (CInt(pageNumber) - 1) * CInt(pageSize)

        ' Build count SQL
        countSql = "SELECT COUNT(1) AS c FROM (" & baseSql & ") AS tcnt"

        ' Build paged SQL using OFFSET/FETCH
        pagedSql = baseSql & " ORDER BY " & orderBy & " OFFSET " & offset & " ROWS FETCH NEXT " & CInt(pageSize) & " ROWS ONLY"

        ' Get total
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = countSql
        cmd.CommandType = 1
        If Not IsMissing(paramsArr) Then
            On Error Resume Next
            Dim ub2, i2, pval2, dtype2, p2
            ub2 = -1
            ub2 = UBound(paramsArr)
            If Err.Number = 0 Then
                For i2 = 0 To ub2
                    pval2 = paramsArr(i2)
                    If IsDate(pval2) Then dtype2 = 7 ElseIf IsNumeric(pval2) Then If InStr(CStr(pval2), ".") > 0 Then dtype2 = 5 Else dtype2 = 3 Else dtype2 = 200
                    Set p2 = cmd.CreateParameter("p" & i2, dtype2, 1, , pval2)
                    cmd.Parameters.Append p2
                Next
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            If Not wrapper Is Nothing Then wrapper.Close
            Set wrapper = Nothing
            Exit Function
        End If
        If Not rs.EOF Then total = rs("c") Else total = 0
        rs.Close
        Set rs = Nothing
        If Not cmd Is Nothing Then Set cmd = Nothing

        ' Now get paged items
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = pagedSql
        cmd.CommandType = 1
        If Not IsMissing(paramsArr) Then
            On Error Resume Next
            ub2 = -1
            ub2 = UBound(paramsArr)
            If Err.Number = 0 Then
                For i2 = 0 To ub2
                    pval2 = paramsArr(i2)
                    If IsDate(pval2) Then dtype2 = 7 ElseIf IsNumeric(pval2) Then If InStr(CStr(pval2), ".") > 0 Then dtype2 = 5 Else dtype2 = 3 Else dtype2 = 200
                    Set p2 = cmd.CreateParameter("p" & i2, dtype2, 1, , pval2)
                    cmd.Parameters.Append p2
                Next
            End If
            On Error GoTo 0
        End If
        On Error Resume Next
        Set rs = cmd.Execute()
        If Err.Number <> 0 Then
            If Not wrapper Is Nothing Then wrapper.Close
            Set wrapper = Nothing
            Exit Function
        End If
        idx = 0
        Do While Not rs.EOF
            If IsNull(items) Then ReDim items(0)
            Set items(idx) = RecordsetToDict(rs)
            idx = idx + 1
            rs.MoveNext
            ReDim Preserve items(idx-1)
        Loop
        rs.Close
        Set rs = Nothing
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing

        Set result = CreateObject("Scripting.Dictionary")
        result.Add "items", items
        result.Add "total", total
        result.Add "page", CInt(pageNumber)
        result.Add "pageSize", CInt(pageSize)
        Set PagedList = result
    End Function

    Public Function Insert(dict)
        Dim keys, vals, k, sql, cmd, wrapper
        Set Insert = Nothing
        If Name = "" Then Exit Function
        If dict Is Nothing Then Exit Function
        keys = "" : vals = ""
        For Each k In dict.Keys
            If keys <> "" Then keys = keys & ","
            keys = keys & k
            If vals <> "" Then vals = vals & ","
            vals = vals & "?"
        Next
        sql = "INSERT INTO " & Name & " (" & keys & ") VALUES (" & vals & ")"
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = sql
        cmd.CommandType = 1
        For Each k In dict.Keys
            cmd.Parameters.Append cmd.CreateParameter("p", 200, 1, 8000, dict.Item(k))
        Next
        On Error Resume Next
        cmd.Execute
        If Err.Number = 0 Then Insert = True Else Insert = False
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
    End Function

    Public Function Update(id, dict)
        Dim setClause, k, sql, cmd, wrapper
        Set Update = Nothing
        If Name = "" Then Exit Function
        setClause = ""
        For Each k In dict.Keys
            If setClause <> "" Then setClause = setClause & ","
            setClause = setClause & k & " = ?"
        Next
        sql = "UPDATE " & Name & " SET " & setClause & " WHERE " & PrimaryKey & " = ?"
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
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
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
    End Function

    Public Function Delete(id)
        Dim sql, cmd, wrapper
        Set Delete = Nothing
        If Name = "" Then Exit Function
        sql = "DELETE FROM " & Name & " WHERE " & PrimaryKey & " = ?"
        Set wrapper = ensureWrapper()
        If wrapper Is Nothing Then Exit Function
        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = dbConnFromObject(wrapper)
        cmd.CommandText = sql
        cmd.CommandType = 1
        cmd.Parameters.Append cmd.CreateParameter("idp", 3, 1, , CInt(id))
        On Error Resume Next
        cmd.Execute
        If Err.Number = 0 Then Delete = True Else Delete = False
        If Not cmd Is Nothing Then Set cmd = Nothing
        If Not wrapper Is Nothing Then wrapper.Close
        Set wrapper = Nothing
    End Function

    Private Function RecordsetToDict(rs)
        Dim d, i
        Set d = CreateObject("Scripting.Dictionary")
        For i = 0 To rs.Fields.Count - 1
            d.Add rs.Fields(i).Name, rs.Fields(i).Value
        Next
        Set RecordsetToDict = d
    End Function

    Private Function dbConnFromObject(obj)
        On Error Resume Next
        Dim t
        t = TypeName(obj)
        If InStr(1, t, "MSSQLConnection", 1) > 0 Then
            Dim c
            Set c = Server.CreateObject("ADODB.Connection")
            c.Open obj.ConnectionString
            Set dbConnFromObject = c
            Exit Function
        Else
            Set dbConnFromObject = obj
            Exit Function
        End If
        Set dbConnFromObject = Nothing
    End Function

End Class
%>
