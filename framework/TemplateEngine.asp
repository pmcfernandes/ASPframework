<!--#include virtual="/" -->
<!-- TemplateEngine.asp
   Simple Classic ASP template engine (Mustache-like features)

   Usage:
     <!--#include file="TemplateEngine.asp" -->
     Dim tpl: Set tpl = New TemplateEngine
     tpl.Load Server.MapPath("/templates/person.tpl")
     tpl.Set "name", "Alice"
     tpl.Set "age", 30
     tpl.SetArray "items", ArrayOfDicts ' where each dict is Scripting.Dictionary
     Response.Write tpl.Render()
     Set tpl = Nothing
-->

<%
Class TemplateEngine
    Private templateText
    Private vars ' Scripting.Dictionary

    Private Sub Class_Initialize()
        Set vars = CreateObject("Scripting.Dictionary")
        templateText = ""
    End Sub

    Public Sub Load(path)
        On Error Resume Next
        Dim fso, t, ts
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FileExists(path) Then
            templateText = ""
            Exit Sub
        End If
        Set t = fso.GetFile(path)
        Set ts = t.OpenAsTextStream(1, -2) ' ForReading, TristateUseDefault
        templateText = ts.ReadAll
        ts.Close
        Set ts = Nothing
        Set t = Nothing
        Set fso = Nothing
    End Sub

    Public Sub Set(key, value)
        If vars.Exists(key) Then vars.Remove key
        vars.Add key, value
    End Sub

    Public Sub SetDict(d)
        ' Accept a Scripting.Dictionary and merge into vars
        Dim k
        For Each k In d.Keys
            Set k = k ' no-op to satisfy VBScript
            If vars.Exists(k) Then vars.Remove k
            vars.Add k, d.Item(k)
        Next
    End Sub

    Public Sub SetArray(key, arr)
        If vars.Exists(key) Then vars.Remove key
        vars.Add key, arr
    End Sub

    Public Function Render()
        Dim out
        out = templateText
        out = RenderSections(out)
        out = RenderVariables(out)
        Render = out
    End Function

    Public Sub RenderToResponse()
        Response.Write Render()
    End Sub

    Private Function RenderVariables(text)
        Dim re, matches, m, name, val
        Set re = New RegExp
        re.Pattern = "\{\{\s*([\.\w]+)\s*\}\}"
        re.IgnoreCase = True
        re.Global = True
        Set matches = re.Execute(text)
        Dim i
        For i = matches.Count - 1 To 0 Step -1
            Set m = matches(i)
            name = m.SubMatches(0)
            val = ResolveVar(name)
            text = Left(text, m.FirstIndex) & EscapeHtml(CStr(val)) & Mid(text, m.FirstIndex + m.Length + 1)
        Next
        RenderVariables = text
    End Function

    Private Function RenderSections(text)
        ' Handle sections like {{#items}}...{{/items}}
        Dim re, matches, m, sectionName, inner, startIdx, endIdx, rendered, arr, i, item
        Set re = New RegExp
        re.Pattern = "\{\{\#\s*(\w+)\s*\}\}([\s\S]*?)\{\{\/\s*\1\s*\}\}"
        re.IgnoreCase = True
        re.Global = True
        Set matches = re.Execute(text)
        For i = matches.Count - 1 To 0 Step -1
            Set m = matches(i)
            sectionName = m.SubMatches(0)
            inner = m.SubMatches(1)
            rendered = ""
            If vars.Exists(sectionName) Then
                arr = vars.Item(sectionName)
                If IsArray(arr) Then
                    Dim idx
                    For Each item In arr
                        Dim tmp
                        tmp = inner
                        If TypeName(item) = "Dictionary" Or TypeName(item) = "Scripting.Dictionary" Then
                            ' Replace {{key}} inside section with values from item
                            Dim k
                            For Each k In item.Keys
                                tmp = Replace(tmp, "{{" & k & "}}", EscapeHtml(CStr(item.Item(k))))
                            Next
                            ' Support {{.}} for scalar items
                            tmp = Replace(tmp, "{{.}}", EscapeHtml(CStr(item)))
                        Else
                            ' scalar
                            tmp = Replace(tmp, "{{.}}", EscapeHtml(CStr(item)))
                        End If
                        ' Render variables inside the tmp
                        tmp = RenderVariables(tmp)
                        rendered = rendered & tmp
                    Next
                End If
            End If
            ' Replace the whole match with rendered
            text = Left(text, m.FirstIndex) & rendered & Mid(text, m.FirstIndex + m.Length + 1)
        Next
        RenderSections = text
    End Function

    Private Function ResolveVar(name)
        ' Support dot notation for nested dicts: person.name
        Dim parts, p, cur, i
        parts = Split(name, ".")
        cur = Null
        If vars.Exists(parts(0)) Then cur = vars.Item(parts(0))
        For i = 1 To UBound(parts)
            If IsNull(cur) Then Exit For
            If TypeName(cur) = "Dictionary" Or TypeName(cur) = "Scripting.Dictionary" Then
                If cur.Exists(parts(i)) Then cur = cur.Item(parts(i)) Else cur = Null
            Else
                cur = Null
            End If
        Next
        If IsNull(cur) Then ResolveVar = "" Else ResolveVar = cur
    End Function

    Private Function EscapeHtml(s)
        If IsNull(s) Then EscapeHtml = "" : Exit Function
        s = Replace(s, "&", "&amp;")
        s = Replace(s, "<", "&lt;")
        s = Replace(s, ">", "&gt;")
        s = Replace(s, """", "&quot;")
        s = Replace(s, "'", "&#39;")
        EscapeHtml = s
    End Function

    Private Sub Class_Terminate()
        On Error Resume Next
        Set vars = Nothing
    End Sub
End Class
%>
