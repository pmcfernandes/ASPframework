<!-- Controller.asp
   Base controller helpers for Classic ASP
   Usage:
     Dim ctrl: Set ctrl = New Controller
     ctrl.Init
     If ctrl.IsPost() Then
       ' handle
     Else
       ctrl.Render "template.tpl"
     End If
-->

<!--#include file="HTTPHelper.asp" -->
<!--#include file="TemplateEngine.asp" -->
<!--#include file="Database.asp" -->
<!--#include file="EnvHelper.asp" -->
<!--#include file="Json.asp" -->

<%
Class Controller
    Public TemplatePath

    Public Sub Init()
        LoadEnv Server.MapPath("/.env")
        TemplatePath = GetEnv("TEMPLATES_DIR", Server.MapPath("/templates"))
    End Sub

    Public Function Param(name, Optional defaultValue)
        Dim v
        v = Request.Form(name)
        If v = "" Then v = Request.QueryString(name)
        If v = "" Then Param = defaultValue Else Param = v
    End Function

    Public Function IsGet()
        IsGet = IsGet()
    End Function

    Public Function IsPost()
        IsPost = IsPost()
    End Function

    Public Function IsPut()
        IsPut = IsPut()
    End Function

    Public Function IsDelete()
        IsDelete = IsDelete()
    End Function

    Public Sub Render(templateFile, Optional dict)
        Dim tpl, path
        Set tpl = New TemplateEngine
        path = Server.MapPath(TemplatePath & "/" & templateFile)
        tpl.Load path
        If Not IsMissing(dict) Then tpl.SetDict dict
        tpl.RenderToResponse
        Set tpl = Nothing
    End Sub

    Public Sub Json(obj)
        Response.ContentType = "application/json"
        Response.Write JsonEncode(obj)
    End Sub

    Public Sub Redirect(url)
        Response.Redirect url
    End Sub

    Public Function CreateDB()
        Dim db, provider, server, name, user, pass, connStr, envConn
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

        Set db = New MSSQLConnection
        db.ConnectionString = connStr
        If db.Connect() Then
            Set CreateDB = db
        Else
            Set CreateDB = Nothing
        End If
    End Function


End Class
%>
