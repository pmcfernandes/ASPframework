<!-- HTTPHelper.asp
   Helpers to detect HTTP method in Classic ASP
   Supports X-HTTP-Method-Override header and a form/query param named _method
-->

<%
Function GetHttpMethod()
    Dim m, override
    m = LCase(Request.ServerVariables("REQUEST_METHOD"))
    ' Check header override
    override = Request.ServerVariables("HTTP_X_HTTP_METHOD_OVERRIDE")
    If override <> "" Then
        m = LCase(override)
    Else
        ' Check form or query param override
        override = LCase(Request.Form("_method"))
        If override = "" Then override = LCase(Request.QueryString("_method"))
        If override <> "" Then m = override
    End If
    GetHttpMethod = UCase(m)
End Function

Function IsGet()
    IsGet = (GetHttpMethod() = "GET")
End Function

Function IsPost()
    IsPost = (GetHttpMethod() = "POST")
End Function

Function IsPut()
    IsPut = (GetHttpMethod() = "PUT")
End Function

Function IsDelete()
    IsDelete = (GetHttpMethod() = "DELETE")
End Function

%>
