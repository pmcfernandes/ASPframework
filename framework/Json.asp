<!-- Json.asp
   Simple JSON encoding for Classic ASP (VBScript)
   Usage:
     Return JsonEncode(obj) ' returns JSON string
-->

<%

  Public Function JsonEncode(obj)
        ' Very small JSON encoder for dictionaries, arrays, scalars
        Dim t
        t = TypeName(obj)
        If IsNull(obj) Then JsonEncode = "null" : Exit Function
        If t = "Dictionary" Or t = "Scripting.Dictionary" Then
            Dim sb, k
            sb = "{"
            For Each k In obj.Keys
                sb = sb & """" & EscapeJson(k) & """:" & JsonEncode(obj.Item(k)) & ","
            Next
            If Right(sb,1) = "," Then sb = Left(sb, Len(sb)-1)
            sb = sb & "}"
            JsonEncode = sb
            Exit Function
        End If
        If IsArray(obj) Then
            Dim arr, i
            arr = obj
            sb = "["
            For i = 0 To UBound(arr)
                sb = sb & JsonEncode(arr(i)) & ","
            Next
            If Right(sb,1) = "," Then sb = Left(sb, Len(sb)-1)
            sb = sb & "]"
            JsonEncode = sb
            Exit Function
        End If
        If VarType(obj) = vbString Then JsonEncode = """" & EscapeJson(obj) & """" : Exit Function
        If VarType(obj) = vbBoolean Then If obj Then JsonEncode = "true" Else JsonEncode = "false" : Exit Function
        If IsNumeric(obj) Then JsonEncode = CStr(obj) : Exit Function
        JsonEncode = """" & EscapeJson(CStr(obj)) & """"
    End Function

    Private Function EscapeJson(s)
        s = Replace(s, "\", "\\")
        s = Replace(s, """", "\"")
        s = Replace(s, vbCrLf, "\n")
        s = Replace(s, vbCr, "\n")
        s = Replace(s, vbLf, "\n")
        EscapeJson = s
    End Function

%>
