<!-- SearchControl.asp
   Render a search form where each search parameter is one line with label and input.
   Usage:
     <!--#include file="SearchControl.asp" -->
     Dim sc: Set sc = New SearchControl
     sc.AddField "first_name", "First name", "text"
     sc.AddField "email", "Email", "text"
     sc.Render "/search_results.asp", "GET"
     ' To read values after submit: sc.GetValue("first_name")
-->

<%
Class SearchControl
    Private fields ' array of dictionaries {name,label,type,attrs}
    Private method
    Private action

    Private Sub Class_Initialize()
        ReDim fields(-1)
        method = "GET"
        action = ""
    End Sub

    Public Sub AddField(name, label, Optional type, Optional attrs, Optional options)
        If IsMissing(type) Then type = "text"
        If IsMissing(attrs) Then attrs = ""
        Dim d
        Set d = CreateObject("Scripting.Dictionary")
        d.Add "name", name
        d.Add "label", label
        d.Add "type", type
        d.Add "attrs", attrs
        If IsMissing(options) Then
            d.Add "options", Empty
        Else
            d.Add "options", options
        End If
        Dim i
        i = UBound(fields) + 1
        ReDim Preserve fields(i)
        Set fields(i) = d
    End Sub

    Public Sub Render(target, Optional httpMethod)
        If Not IsMissing(httpMethod) Then method = UCase(httpMethod)
        action = target
        Dim i, f, html
        html = "<form method='" & Server.HTMLEncode(method) & "' action='" & Server.HTMLEncode(action) & "'>"
        For i = 0 To UBound(fields)
            Set f = fields(i)
            html = html & RenderField(f)
        Next
        html = html & "<div class=\"search-row\"><input type='submit' value='Search' /></div>"
        html = html & "</form>"
        Response.Write html
    End Sub

    Public Function GetValue(name)
        Dim v
        v = Request.Form(name)
        If v = "" Then v = Request.QueryString(name)
        GetValue = v
    End Function

    ' Build a parameterized WHERE clause and a params array based on non-empty fields
    ' Returns a Scripting.Dictionary with keys: "where" (string, e.g. "col1 = ? AND col2 LIKE ?") and "params" (array)
    Public Function BuildWhereFromFields()
        Dim dict, clauses(), paramsArr, pcount, i, f, nm, t, v, clause
        Set dict = CreateObject("Scripting.Dictionary")
        ReDim paramsArr(-1)
        pcount = 0
        For i = 0 To UBound(fields)
            Set f = fields(i)
            nm = f.Item("name")
            t = LCase(f.Item("type"))
            v = GetValue(nm)
            If Not (v = "" Or IsNull(v)) Then
                ' Determine operator based on type or options
                If t = "text" Or t = "textarea" Then
                    clause = nm & " LIKE ?"
                    ReDim Preserve paramsArr(pcount)
                    paramsArr(pcount) = "%" & v & "%"
                ElseIf t = "select" Or t = "date" Or t = "number" Or t = "integer" Then
                    clause = nm & " = ?"
                    ReDim Preserve paramsArr(pcount)
                    paramsArr(pcount) = v
                ElseIf t = "checkbox" Then
                    clause = nm & " = ?"
                    ReDim Preserve paramsArr(pcount)
                    paramsArr(pcount) = v
                Else
                    clause = nm & " = ?"
                    ReDim Preserve paramsArr(pcount)
                    paramsArr(pcount) = v
                End If
                If pcount = 0 Then
                    dict.Add "where", clause
                Else
                    dict.Item("where") = dict.Item("where") & " OR " & clause
                End If
                pcount = pcount + 1
            End If
        Next
        If pcount = 0 Then
            dict.Add "where", ""
        End If
        dict.Add "params", paramsArr
        Set BuildWhereFromFields = dict
    End Function

    Private Function RenderField(f)
        Dim t, nm, lbl, attrs, opts, v, out
        t = LCase(f.Item("type"))
        nm = f.Item("name")
        lbl = f.Item("label")
        attrs = f.Item("attrs")
        opts = f.Item("options")
        v = GetValue(nm)
        out = ""
        out = out & "<div class=\"search-row\">"
        out = out & "<label for='" & Server.HTMLEncode(nm) & "'>" & Server.HTMLEncode(lbl) & "</label>"
        If t = "select" Then
            out = out & "<select name='" & Server.HTMLEncode(nm) & "' " & attrs & ">"
            ' options can be array, dictionary, or comma-separated string
            If Not IsEmpty(opts) Then
                Dim k
                If IsArray(opts) Then
                    For Each k In opts
                        out = out & "<option value='" & Server.HTMLEncode(CStr(k)) & "'"
                        If CStr(k) = CStr(v) Then out = out & " selected"
                        out = out & ">" & Server.HTMLEncode(CStr(k)) & "</option>"
                    Next
                ElseIf TypeName(opts) = "Dictionary" Or TypeName(opts) = "Scripting.Dictionary" Then
                    For Each k In opts.Keys
                        out = out & "<option value='" & Server.HTMLEncode(CStr(k)) & "'"
                        If CStr(k) = CStr(v) Then out = out & " selected"
                        out = out & ">" & Server.HTMLEncode(CStr(opts.Item(k))) & "</option>"
                    Next
                Else
                    ' assume comma separated string
                    Dim pieces
                    pieces = Split(opts, ",")
                    Dim pi
                    For pi = 0 To UBound(pieces)
                        Dim val
                        val = Trim(pieces(pi))
                        out = out & "<option value='" & Server.HTMLEncode(val) & "'"
                        If CStr(val) = CStr(v) Then out = out & " selected"
                        out = out & ">" & Server.HTMLEncode(val) & "</option>"
                    Next
                End If
            End If
            out = out & "</select>"
        ElseIf t = "checkbox" Then
            Dim checkedVal
            checkedVal = "1"
            ' if options provided, take first option as checked value
            If Not IsEmpty(opts) Then
                If IsArray(opts) Then checkedVal = opts(0) Else If TypeName(opts) = "String" Then checkedVal = opts
            End If
            out = out & "<input type='checkbox' name='" & Server.HTMLEncode(nm) & "' value='" & Server.HTMLEncode(CStr(checkedVal)) & "'"
            If CStr(v) = CStr(checkedVal) Or (v <> "" And CStr(checkedVal) = "1" And CStr(v) <> "") Then out = out & " checked"
            out = out & " " & attrs & " />"
        ElseIf t = "textarea" Then
            out = out & "<textarea name='" & Server.HTMLEncode(nm) & "' " & attrs & ">" & Server.HTMLEncode(CStr(v)) & "</textarea>"
        ElseIf t = "date" Then
            ' normalize value to yyyy-mm-dd if possible
            Dim dv
            dv = ""
            If v <> "" Then
                On Error Resume Next
                dv = Year(CDate(v)) & "-" & Right("0" & Month(CDate(v)),2) & "-" & Right("0" & Day(CDate(v)),2)
                If Err.Number <> 0 Then dv = v
                Err.Clear
            End If
            out = out & "<input type='date' name='" & Server.HTMLEncode(nm) & "' value='" & Server.HTMLEncode(dv) & "' " & attrs & " />"
        Else
            out = out & "<input type='" & Server.HTMLEncode(t) & "' name='" & Server.HTMLEncode(nm) & "' value='" & Server.HTMLEncode(CStr(v)) & "' " & attrs & " />"
        End If
        out = out & "</div>"
        RenderField = out
    End Function

    Public Sub Clear()
        ReDim fields(-1)
    End Sub

    ' Prepare an ADODB.Command for a base SQL string using the filters from this SearchControl.
    ' baseSql: e.g. "SELECT * FROM Users"
    ' connectionOrWrapper: optional. Can be an ADODB.Connection or a MSSQLConnection wrapper (uses its ConnectionString to open a connection).
    ' Returns an ADODB.Command with CommandText set (WHERE clause appended if any) and Parameters appended in-order.
    Public Function PrepareCommand(baseSql, Optional connectionOrWrapper)
        Dim whereDict, whereClause, fullSql
        Set whereDict = BuildWhereFromFields()
        whereClause = ""
        If whereDict.Exists("where") Then whereClause = whereDict.Item("where")
        fullSql = baseSql
        If Len(Trim(CStr(whereClause))) > 0 Then
            fullSql = fullSql & " WHERE " & whereClause
        End If

        Dim cmd
        Set cmd = Server.CreateObject("ADODB.Command")
        cmd.CommandText = fullSql
        cmd.CommandType = 1 ' adCmdText

        ' Attach or open a connection if provided
        If Not IsMissing(connectionOrWrapper) Then
            On Error Resume Next
            Dim tname
            tname = TypeName(connectionOrWrapper)
            If InStr(1, tname, "MSSQLConnection", 1) > 0 Then
                ' Open a fresh ADODB.Connection using the wrapper's ConnectionString (wrapper exposes ConnectionString property)
                Dim tmpCon
                Set tmpCon = Server.CreateObject("ADODB.Connection")
                tmpCon.Open connectionOrWrapper.ConnectionString
                If Err.Number = 0 Then
                    Set cmd.ActiveConnection = tmpCon
                Else
                    Err.Clear
                End If
            Else
                ' Assume it's an ADODB.Connection or similar
                On Error Resume Next
                Set cmd.ActiveConnection = connectionOrWrapper
                Err.Clear
            End If
            On Error GoTo 0
        End If

        ' Append parameters in the same order produced by BuildWhereFromFields
        Dim paramsArr
        paramsArr = whereDict.Item("params")
        On Error Resume Next
        Dim ub
        ub = -1
        ub = UBound(paramsArr)
        If Err.Number = 0 Then
            Dim j, pval, dtype, psize, p
            For j = 0 To ub
                pval = paramsArr(j)
                ' Heuristic for data type
                If IsDate(pval) Then
                    dtype = 7 ' adDate
                    psize = 0
                ElseIf IsNumeric(pval) Then
                    If InStr(CStr(pval), ".") > 0 Then
                        dtype = 5 ' adDouble
                    Else
                        dtype = 3 ' adInteger
                    End If
                    psize = 0
                ElseIf VarType(pval) = 11 Then
                    dtype = 11 ' adBoolean
                    psize = 0
                Else
                    dtype = 200 ' adVarChar
                    psize = Len(CStr(pval))
                    If psize = 0 Then psize = 255
                End If

                Set p = cmd.CreateParameter("p" & j, dtype, 1, psize, pval)
                cmd.Parameters.Append p
            Next
        End If
        Err.Clear
        On Error GoTo 0

        Set PrepareCommand = cmd
    End Function

    Private Sub Class_Terminate()
        On Error Resume Next
        Dim i
        For i = 0 To UBound(fields)
            Set fields(i) = Nothing
        Next
    End Sub
End Class
%>
