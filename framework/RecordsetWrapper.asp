<!--#include virtual="/" -->
<!-- RecordsetWrapper.asp
   RecordsetWrapper class for Classic ASP (VBScript)

   Usage:
     <!--#include file="RecordsetWrapper.asp" -->
     Dim rsw
     Set rsw = New RecordsetWrapper
     rsw.Init rs ' pass an ADODB.Recordset
     Response.Write rsw.ToJSON()
     rsw.Close
     Set rsw = Nothing
-->

<%
Class RecordsetWrapper
    Private rs

    Public Sub Init(recordset)
        Set rs = recordset
    End Sub

    Public Function FieldExists(fieldName)
        On Error Resume Next
        Dim f
        FieldExists = False
        If rs Is Nothing Then Exit Function
        f = rs.Fields(fieldName)
        If Err.Number = 0 Then FieldExists = True
        Err.Clear
    End Function

    Public Function GetValue(fieldName, Optional defaultValue)
        On Error Resume Next
        If rs Is Nothing Then
            GetValue = defaultValue
            Exit Function
        End If

        If FieldExists(fieldName) Then
            GetValue = rs(fieldName)
            If IsNull(GetValue) Then GetValue = defaultValue
        Else
            GetValue = defaultValue
        End If
    End Function

    Public Function ToArrayOfDicts()
        ' Returns a VBScript array of Scripting.Dictionary objects representing rows
        Dim arr(), count, dict
        If rs Is Nothing Then
            ReDim arr(-1)
            ToArrayOfDicts = arr
            Exit Function
        End If

        rs.MoveFirst
        count = 0
        Do While Not rs.EOF
            Set dict = CreateObject("Scripting.Dictionary")
            Dim i
            For i = 0 To rs.Fields.Count - 1
                dict.Add rs.Fields(i).Name, rs.Fields(i).Value
            Next
            ReDim Preserve arr(count)
            Set arr(count) = dict
            count = count + 1
            rs.MoveNext
        Loop

        ToArrayOfDicts = arr
    End Function

    Public Function ToJSON()
        ' Basic JSON serialization. Not fully RFC-complete but suitable for simple use.
        Dim arr, i, j, row, sb, fldName, val
        If rs Is Nothing Then
            ToJSON = "[]"
            Exit Function
        End If

        arr = ToArrayOfDicts()
        sb = "["
        For i = 0 To UBound(arr)
            Set row = arr(i)
            sb = sb & "{"
            j = 0
            For Each fldName In row.Keys
                val = row.Item(fldName)
                sb = sb & """" & EscapeJsonString(fldName) & """:" & JsonValue(val)
                If j < row.Count - 1 Then sb = sb & ","
                j = j + 1
            Next
            sb = sb & "}"
            If i < UBound(arr) Then sb = sb & ","
        Next
        sb = sb & "]"
        ToJSON = sb
    End Function

    Private Function JsonValue(v)
        If IsNull(v) Then
            JsonValue = "null"
        ElseIf VarType(v) = vbString Then
            JsonValue = """" & EscapeJsonString(v) & """"
        ElseIf VarType(v) = vbBoolean Then
            If v Then JsonValue = "true" Else JsonValue = "false"
        ElseIf IsNumeric(v) Then
            JsonValue = CStr(v)
        ElseIf IsDate(v) Then
            JsonValue = """" & EscapeJsonString(FormatDateTime(v, vbLongDate) & " " & FormatDateTime(v, vbLongTime)) & """"
        Else
            JsonValue = """" & EscapeJsonString(CStr(v)) & """"
        End If
    End Function

    Private Function EscapeJsonString(s)
        If IsNull(s) Then
            EscapeJsonString = ""
            Exit Function
        End If
        s = Replace(s, "\", "\\")
        s = Replace(s, """", "\"")
        s = Replace(s, vbCrLf, "\n")
        s = Replace(s, vbCr, "\n")
        s = Replace(s, vbLf, "\n")
        EscapeJsonString = s
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
