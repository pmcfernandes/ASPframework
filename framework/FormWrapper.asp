<!-- FormWrapper.asp
   Classic ASP helper to render HTML form controls prefilled from an ADODB.Recordset

   Usage:
     Set fb = New FormWrapper
     fb.Init rs ' pass an opened ADODB.Recordset positioned at the row to bind
     Response.Write fb.InputText("FirstName", "class=\"form-control\"")
     Response.Write fb.TextArea("Notes", "rows=4 cols=40")
     Response.Write fb.SelectFromRS("CountryID", rsCountries, "ID", "Name")
     fb.Close
     Set fb = Nothing
-->

<%
Class FormWrapper
    Private rs

    Public Sub Init(recordset)
        Set rs = recordset
    End Sub

    Private Function ExistsField(fieldName)
        On Error Resume Next
        Dim f
        ExistsField = False
        If rs Is Nothing Then Exit Function
        f = rs.Fields(fieldName)
        If Err.Number = 0 Then ExistsField = True
        Err.Clear
    End Function

    Private Function RawValue(fieldName)
        On Error Resume Next
        Dim v
        If rs Is Nothing Then
            RawValue = Null
            Exit Function
        End If
        If ExistsField(fieldName) Then
            v = rs(fieldName).Value
        Else
            v = Null
        End If
        RawValue = v
    End Function

    Public Function FieldValue(fieldName, Optional defaultValue)
        Dim v
        v = RawValue(fieldName)
        If IsNull(v) Then
            FieldValue = CStr(defaultValue)
        Else
            FieldValue = Server.HTMLEncode(CStr(v))
        End If
    End Function

    Public Function InputText(name, Optional attrs, Optional valueOverride)
        If IsMissing(attrs) Then attrs = ""
        Dim val
        If Not IsMissing(valueOverride) Then
            val = Server.HTMLEncode(CStr(valueOverride))
        Else
            val = FieldValue(name, "")
        End If
        InputText = "<input type=\"text\" name=\"" & Server.HTMLEncode(name) & "\" value=\"" & val & "\" " & attrs & " />"
    End Function

    Public Function Hidden(name, Optional valueOverride)
        Dim val
        If Not IsMissing(valueOverride) Then
            val = Server.HTMLEncode(CStr(valueOverride))
        Else
            val = FieldValue(name, "")
        End If
        Hidden = "<input type=\"hidden\" name=\"" & Server.HTMLEncode(name) & "\" value=\"" & val & "\" />"
    End Function

    Public Function TextArea(name, Optional attrs, Optional valueOverride)
        If IsMissing(attrs) Then attrs = ""
        Dim val
        If Not IsMissing(valueOverride) Then
            val = Server.HTMLEncode(CStr(valueOverride))
        Else
            val = FieldValue(name, "")
        End If
        TextArea = "<textarea name=\"" & Server.HTMLEncode(name) & "\" " & attrs & ">" & val & "</textarea>"
    End Function

    Public Function Checkbox(name, Optional checkedValue, Optional attrs)
        ' If field equals checkedValue (or truthy when checkedValue not provided) then checkbox is checked
        If IsMissing(attrs) Then attrs = ""
        Dim val, checkedAttr
        val = RawValue(name)
        checkedAttr = ""
        If Not IsMissing(checkedValue) Then
            If Not IsNull(val) And CStr(val) = CStr(checkedValue) Then checkedAttr = " checked"
        Else
            If Not (IsNull(val) Or CStr(val) = "" ) Then checkedAttr = " checked"
        End If
        Checkbox = "<input type=\"checkbox\" name=\"" & Server.HTMLEncode(name) & "\" value=\"1\"" & checkedAttr & " " & attrs & " />"
    End Function

    Public Function SelectFromRS(name, optionsRS, valueField, textField, Optional attrs, Optional includeBlank)
        ' optionsRS: recordset with rows to build option list
        If IsMissing(attrs) Then attrs = ""
        If IsMissing(includeBlank) Then includeBlank = False
        Dim sb, currentVal
        currentVal = RawValue(name)
        sb = "<select name=\"" & Server.HTMLEncode(name) & "\" " & attrs & ">"
        If includeBlank Then sb = sb & "<option value=\"\">" & "" & "</option>"
        If Not optionsRS Is Nothing Then
            optionsRS.MoveFirst
            Do While Not optionsRS.EOF
                Dim optVal, optText, sel
                optVal = optionsRS(valueField)
                optText = optionsRS(textField)
                sel = ""
                If Not IsNull(currentVal) Then
                    If CStr(currentVal) = CStr(optVal) Then sel = " selected"
                End If
                sb = sb & "<option value=\"" & Server.HTMLEncode(CStr(optVal)) & "\"" & sel & ">" & Server.HTMLEncode(CStr(optText)) & "</option>"
                optionsRS.MoveNext
            Loop
        End If
        sb = sb & "</select>"
        SelectFromRS = sb
    End Function

    Public Sub Close()
        On Error Resume Next
        If Not rs Is Nothing Then
            If rs.State = 1 Then rs.Close
            Set rs = Nothing
        End If
    End Sub
End Class
%>
