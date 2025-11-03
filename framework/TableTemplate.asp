<!-- TableTemplate.asp
   Helper include that renders an HTML table from headers and rows.

   Usage:
     ' headers: array of field names OR array of dictionaries {"field": "colName", "label": "Column Label"}
     ' rows: array of Scripting.Dictionary objects (one per row)
     Dim opts: Set opts = CreateObject("Scripting.Dictionary")
     opts.Add "classes", "table table-striped"
     opts.Add "id", "usersTable"
     opts.Add "showActions", True
     opts.Add "actionsBaseUrl", "/edit_user.asp"
     Call RenderTable(headers, rows, opts)
-->

<%
<!--#include file="Pagination.asp" -->

Sub RenderTable(headers, rows, Optional options)
    Dim classes, idAttr, showActions, actionsBaseUrl, actionQueryParam, colCount
    classes = "table"
    idAttr = ""
    showActions = False
    actionsBaseUrl = ""
    actionQueryParam = "id"

    If Not IsMissing(options) Then
        On Error Resume Next
        If options.Exists("classes") Then classes = options.Item("classes")
        If options.Exists("id") Then idAttr = options.Item("id")
        If options.Exists("showActions") Then showActions = CBool(options.Item("showActions"))
        If options.Exists("actionsBaseUrl") Then actionsBaseUrl = options.Item("actionsBaseUrl")
        If options.Exists("formatters") Then Set formatters = options.Item("formatters") Else Set formatters = Nothing
        If options.Exists("links") Then Set links = options.Item("links") Else Set links = Nothing
        If options.Exists("sortable") Then Set sortable = options.Item("sortable") Else Set sortable = Nothing
        If options.Exists("expandable") Then Set expandable = options.Item("expandable") Else Set expandable = Nothing
        If options.Exists("actionQueryParam") Then actionQueryParam = options.Item("actionQueryParam")
        Err.Clear
        On Error GoTo 0
    End If

    ' Determine column headers
    Dim hdrs, i, h, fld, lbl
    hdrs = headers
    colCount = 0
    If IsArray(hdrs) Then
        For i = 0 To UBound(hdrs)
            If TypeName(hdrs(i)) = "Dictionary" Or TypeName(hdrs(i)) = "Scripting.Dictionary" Then
                fld = hdrs(i).Item("field")
            Else
                fld = CStr(hdrs(i))
            End If
            colCount = colCount + 1
        Next
    Else
        ' If headers not array, try to infer from first row
        If Not IsNull(rows) Then
            If UBound(rows) >= 0 Then
                Dim firstRow
                Set firstRow = rows(0)
                Dim keyArr()
                ReDim keyArr(firstRow.Count-1)
                Dim k, p
                p = 0
                For Each k In firstRow.Keys
                    keyArr(p) = k
                    p = p + 1
                Next
                hdrs = keyArr
                colCount = UBound(hdrs) + 1
            End If
        End If
    End If

    Response.Write "<div class=\"table-responsive\">"
    Dim tableOpen
    tableOpen = "<table class='" & Server.HTMLEncode(classes) & "'"
    If idAttr <> "" Then tableOpen = tableOpen & " id='" & Server.HTMLEncode(idAttr) & "'"
    tableOpen = tableOpen & ">"
    Response.Write tableOpen

    ' Render header (with optional sortable links)
    Response.Write "<thead><tr>"
    If IsArray(hdrs) Then
        For i = 0 To UBound(hdrs)
            h = hdrs(i)
            If TypeName(h) = "Dictionary" Or TypeName(h) = "Scripting.Dictionary" Then
                fld = h.Item("field")
                If h.Exists("label") Then lbl = h.Item("label") Else lbl = fld
            Else
                fld = CStr(h)
                lbl = fld
            End If
            Dim headerHtml
            headerHtml = Server.HTMLEncode(lbl)
                    If Not IsEmpty(sortable) And Not sortable Is Nothing Then
                On Error Resume Next
                If sortable.Exists(fld) Then
                    Dim sortParam, currentOrder, newOrder, sortUrl
                    sortParam = sortable.Item("paramName")
                    If sortParam = "" Then sortParam = "orderBy"
                    currentOrder = Request.QueryString(sortParam)
                    newOrder = fld & " ASC"
                    If InStr(currentOrder, fld) > 0 Then
                        If InStr(UCase(currentOrder), "DESC") > 0 Then
                            newOrder = fld & " ASC"
                        Else
                            newOrder = fld & " DESC"
                        End If
                    End If
                    		    sortUrl = Request.ServerVariables("SCRIPT_NAME") & "?" & Server.URLEncode(sortParam) & "=" & Server.URLEncode(newOrder)
                    		    headerHtml = "<a href='" & Server.HTMLEncode(sortUrl) & "' data-field='" & Server.HTMLEncode(fld) & "'>" & headerHtml & "</a>"
                End If
                Err.Clear
                On Error GoTo 0
            End If
            		Response.Write "<th>" & headerHtml & "</th>"
        Next
    End If
    If showActions Then Response.Write "<th>Actions</th>"
    Response.Write "</tr></thead>"

    ' Render body
    Response.Write "<tbody>"
    If IsNull(rows) Then
        Response.Write "<tr><td colspan='" & (colCount + (IIf(showActions,1,0))) & "'>No results</td></tr>"
    Else
        Dim r, colIndex, cellVal
        For i = 0 To UBound(rows)
            Set r = rows(i)
            ' Build the opening TR, add has-expand attribute if expandable provided
            Dim trOpen
            trOpen = "<tr"
            If Not IsEmpty(expandable) And Not expandable Is Nothing Then
                On Error Resume Next
                If expandable.Exists("template") Then
                    trOpen = trOpen & " class='has-expand' data-row-index='" & CStr(i) & "' style='cursor:pointer;'"
                End If
                Err.Clear
                On Error GoTo 0
            End If
            trOpen = trOpen & ">"
            Response.Write trOpen
            If IsArray(hdrs) Then
                For colIndex = 0 To UBound(hdrs)
                    h = hdrs(colIndex)
                    If TypeName(h) = "Dictionary" Or TypeName(h) = "Scripting.Dictionary" Then
                        fld = h.Item("field")
                    Else
                        fld = CStr(h)
                    End If
                    cellVal = ""
                    On Error Resume Next
                    If r.Exists(fld) Then
                        cellVal = r.Item(fld)
                    Else
                        cellVal = ""
                    End If
                    Err.Clear
                    On Error GoTo 0

                    ' Apply formatter if provided
                    Dim outVal
                    outVal = cellVal
                    If Not IsEmpty(formatters) And Not formatters Is Nothing Then
                        On Error Resume Next
                        If formatters.Exists(fld) Then
                            Dim fmt
                            fmt = formatters.Item(fld)
                            Select Case LCase(fmt)
                                Case "date"
                                    If outVal <> "" Then outVal = Year(CDate(outVal)) & "-" & Right("0" & Month(CDate(outVal)),2) & "-" & Right("0" & Day(CDate(outVal)),2)
                                Case "datetime"
                                    If outVal <> "" Then outVal = CStr(outVal) ' leave as-is or format as needed
                                Case "currency"
                                    If IsNumeric(outVal) Then outVal = FormatNumber(CDbl(outVal),2)
                                Case Else
                                    ' allow a VBScript function name for custom formatter
                                    If IsObject(fmt) Then
                                        On Error Resume Next
                                        outVal = fmt(outVal, r)
                                    ElseIf IsString(fmt) Then
                                        ' not calling by name to avoid Eval; user can pass function object
                                    End If
                            End Select
                        End If
                        Err.Clear
                        On Error GoTo 0
                    End If

                    ' Apply link template if provided
                    Dim cellHtml
                    cellHtml = Server.HTMLEncode(CStr(outVal))
                    If Not IsEmpty(links) And Not links Is Nothing Then
                        On Error Resume Next
                        If links.Exists(fld) Then
                            Dim tpl, linkUrl
                            tpl = links.Item(fld) ' e.g. "/users/view?id={id}" or a function
                            If IsObject(tpl) Then
                                linkUrl = tpl(outVal, r)
                            Else
                                linkUrl = tpl
                                ' replace {field} tokens
                                Dim key
                                For Each key In r.Keys
                                    linkUrl = Replace(linkUrl, "{" & key & "}", Server.URLEncode(CStr(r.Item(key))))
                                Next
                            End If
                            If linkUrl <> "" Then cellHtml = "<a href='" & Server.HTMLEncode(linkUrl) & "'>" & cellHtml & "</a>"
                        End If
                        Err.Clear
                        On Error GoTo 0
                    End If

                    Response.Write "<td>" & cellHtml & "</td>"
                Next
            End If
            If showActions Then
                Dim pkVal, actionUrl
                pkVal = ""
                On Error Resume Next
                If r.Exists(actionQueryParam) Then
                    pkVal = r.Item(actionQueryParam)
                ElseIf r.Exists(PrimaryKey) Then
                    pkVal = r.Item(PrimaryKey)
                ElseIf r.Exists("id") Then
                    pkVal = r.Item("id")
                End If
                Err.Clear
                On Error GoTo 0

                actionUrl = ""
                If actionsBaseUrl <> "" Then
                    actionUrl = actionsBaseUrl & "?" & Server.URLEncode(actionQueryParam) & "=" & Server.URLEncode(CStr(pkVal))
                End If
                Response.Write "<td>"
                If actionUrl <> "" Then
                    Response.Write "<a class='action-view' href='" & Server.HTMLEncode(actionUrl) & "'>View</a> "
                    Response.Write "<a class='action-edit' href='" & Server.HTMLEncode(actionsBaseUrl & "/edit.asp?id=" & Server.URLEncode(CStr(pkVal))) & "'>Edit</a> "
                Else
                    Response.Write "&nbsp;"
                End If
                Response.Write "</td>"
            End If
            ' Expandable row support: render a following hidden row with extra content if expandable option provided
            If Not IsEmpty(expandable) And Not expandable Is Nothing Then
                On Error Resume Next
                If expandable.Exists("template") Then
                    Dim expandHtml
                    Dim tplExp
                    tplExp = expandable.Item("template")
                    If IsObject(tplExp) Then
                        expandHtml = tplExp(r)
                    Else
                        ' treat as format string with {field} tokens
                        expandHtml = tplExp
                        Dim ak
                        For Each ak In r.Keys
                            expandHtml = Replace(expandHtml, "{" & ak & "}", Server.HTMLEncode(CStr(r.Item(ak))))
                        Next
                    End If
                    Response.Write "<tr class='expand-row' style='display:none;'><td colspan='" & (colCount + (IIf(showActions,1,0))) & "'>" & expandHtml & "</td></tr>"
                End If
                Err.Clear
                On Error GoTo 0
            End If
            Response.Write "</tr>"
        Next
    End If
    Response.Write "</tbody>"
    Response.Write "</table>"

    ' Small JS/CSS to enable expandable rows and sort icons
    Response.Write "<style>.expand-row{background:#f9f9f9;} th a{color:inherit;text-decoration:none;} .sort-icon{margin-left:6px;font-size:0.8em}></style>"
    Response.Write "<script>document.addEventListener('DOMContentLoaded',function(){try{var table=document.getElementById('" & Server.HTMLEncode(idAttr) & "'); if(!table){table=document.querySelector('.table-responsive table');} if(table){var orderParam='orderBy', params=new URLSearchParams(location.search), order=params.get(orderParam)||''; // set sort icons
    var anchors=table.querySelectorAll('thead th a[data-field]'); anchors.forEach(function(a){var f=a.getAttribute('data-field'); if(order && order.indexOf(f)===0){var dir=(order.toUpperCase().indexOf('DESC')>-1)?'▼':'▲'; var span=document.createElement('span'); span.className='sort-icon'; span.textContent=dir; a.appendChild(span);} });
    // expandable rows toggle
    var rows=table.querySelectorAll('tbody tr.has-expand'); rows.forEach(function(r){r.addEventListener('click',function(e){if(e.target && (e.target.tagName==='A' || e.target.closest('a'))) return; var next=r.nextElementSibling; if(next && next.classList.contains('expand-row')){ if(next.style.display==='none') next.style.display='table-row'; else next.style.display='none'; }}); }); }
    }catch(e){/* ignore */}});</script>"
    Response.Write "</div>"
End Sub

' Convenience: RenderTableWithPager(tableObj, baseSql, paramsArr, Optional headers, Optional options)
' Runs tableObj.PagedList using request params (page/pageSize/orderBy) and renders the table and pagination
Sub RenderTableWithPager(tableObj, baseSql, paramsArr, Optional headers, Optional options)
    If IsMissing(options) Then Set options = Nothing
    Dim pager
    Set pager = GetPagerParams(20)
    Dim pageSize, pageNumber, orderBy
    pageSize = pager.Item("pageSize")
    pageNumber = pager.Item("page")
    orderBy = Request.QueryString("orderBy")
    If Trim(orderBy) = "" Then orderBy = ""

    Dim res
    Set res = tableObj.PagedList(baseSql, paramsArr, orderBy, pageSize, pageNumber)
    Dim items
    Set items = res.Item("items")
    Call RenderTable(headers, items, options)
    Call RenderPagination(res.Item("total"), res.Item("page"), res.Item("pageSize"), options)
End Sub
%>
