<!-- Pagination.asp
   Standalone pagination include for Classic ASP
   Usage:
     <!--#include file="framework/Pagination.asp" -->
     Call RenderPagination(total, currentPage, pageSize)
-->

<%
' RenderPagination(total, currentPage, pageSize, Optional options)
' Renders First, Prev, numbered pages (max 10), Next, Last and a pageSize selector (10/20/50/100)
Sub RenderPagination(total, currentPage, pageSize, Optional options)
    Dim opts
    If IsMissing(options) Then
        Set opts = Nothing
    Else
        Set opts = options
    End If

    If Not IsNumeric(total) Then total = 0
    If Not IsNumeric(currentPage) Then currentPage = 1
    If Not IsNumeric(pageSize) Then pageSize = 20

    Dim totalPages
    totalPages = Int((CInt(total) + CInt(pageSize) - 1) / CInt(pageSize))
    If totalPages < 1 Then totalPages = 1

    Dim maxButtons
    maxButtons = 10
    Dim startPage, endPage
    If totalPages <= maxButtons Then
        startPage = 1
        endPage = totalPages
    Else
        startPage = currentPage - (maxButtons \ 2)
        If startPage < 1 Then startPage = 1
        endPage = startPage + maxButtons - 1
        If endPage > totalPages Then
            endPage = totalPages
            startPage = endPage - maxButtons + 1
        End If
    End If

    Dim qsRaw
    qsRaw = Request.ServerVariables("QUERY_STRING")

    Function MergeQueryString(newParams)
        Dim dict, pairs, i, p, k, v, key, val, parts, outParts()
        Set dict = CreateObject("Scripting.Dictionary")
        If Len(Trim(qsRaw)) > 0 Then
            pairs = Split(qsRaw, "&")
            For i = 0 To UBound(pairs)
                p = pairs(i)
                If InStr(p, "=") > 0 Then
                    key = Left(p, InStr(p, "=")-1)
                    val = Mid(p, InStr(p, "=")+1)
                Else
                    key = p
                    val = ""
                End If
                If Not dict.Exists(key) Then dict.Add key, val Else dict.Item(key) = val
            Next
        End If
        For Each k In newParams.Keys
            dict.Item(k) = newParams.Item(k)
        Next
        ReDim outParts(dict.Count-1)
        i = 0
        For Each k In dict.Keys
            outParts(i) = k & "=" & Server.URLEncode(CStr(dict.Item(k)))
            i = i + 1
        Next
        MergeQueryString = Join(outParts, "&")
    End Function

    Function BuildPageUrl(targetPage, targetPageSize)
        Dim np
        Set np = CreateObject("Scripting.Dictionary")
        np.Add "page", CStr(targetPage)
        np.Add "pageSize", CStr(targetPageSize)
        Dim qs
        qs = MergeQueryString(np)
        BuildPageUrl = Request.ServerVariables("SCRIPT_NAME") & "?" & qs
    End Function

    Response.Write "<nav class='pagination-container' aria-label='Pagination'><ul class='pagination'>"

    If CInt(currentPage) > 1 Then
        Response.Write "<li class='page-item'><a class='page-link' href='" & Server.HTMLEncode(BuildPageUrl(1, pageSize)) & "'>&laquo; First</a></li>"
    Else
        Response.Write "<li class='page-item disabled'><span class='page-link'>&laquo; First</span></li>"
    End If

    If CInt(currentPage) > 1 Then
        Response.Write "<li class='page-item'><a class='page-link' href='" & Server.HTMLEncode(BuildPageUrl(CInt(currentPage)-1, pageSize)) & "'>&lsaquo; Prev</a></li>"
    Else
        Response.Write "<li class='page-item disabled'><span class='page-link'>&lsaquo; Prev</span></li>"
    End If

    Dim pg
    For pg = startPage To endPage
        If pg = CInt(currentPage) Then
            Response.Write "<li class='page-item active'><span class='page-link'>" & CStr(pg) & "</span></li>"
        Else
            Response.Write "<li class='page-item'><a class='page-link' href='" & Server.HTMLEncode(BuildPageUrl(pg, pageSize)) & "'>" & CStr(pg) & "</a></li>"
        End If
    Next

    If CInt(currentPage) < CInt(totalPages) Then
        Response.Write "<li class='page-item'><a class='page-link' href='" & Server.HTMLEncode(BuildPageUrl(CInt(currentPage)+1, pageSize)) & "'>Next &rsaquo;</a></li>"
    Else
        Response.Write "<li class='page-item disabled'><span class='page-link'>Next &rsaquo;</span></li>"
    End If

    If CInt(currentPage) < CInt(totalPages) Then
        Response.Write "<li class='page-item'><a class='page-link' href='" & Server.HTMLEncode(BuildPageUrl(totalPages, pageSize)) & "'>Last &raquo;</a></li>"
    Else
        Response.Write "<li class='page-item disabled'><span class='page-link'>Last &raquo;</span></li>"
    End If

    Response.Write "</ul>"

    Dim sizes, s, selHtml
    sizes = Array(10,20,50,100)
    selHtml = "<label for='pageSizeSelect'>Per page:</label> <select id='pageSizeSelect' class='page-size-select' onchange='(function(){var v=this.value;var u=new URL(location.href);u.searchParams.set(""pageSize"",v);u.searchParams.set(""page"",1);location.href=u.toString();}).call(this)'>"
    For Each s In sizes
        If CInt(s) = CInt(pageSize) Then
            selHtml = selHtml & "<option value='" & CStr(s) & "' selected>" & CStr(s) & "</option>"
        Else
            selHtml = selHtml & "<option value='" & CStr(s) & "'>" & CStr(s) & "</option>"
        End If
    Next
    selHtml = selHtml & "</select>"
    Response.Write selHtml

    Response.Write "</nav>"
End Sub

' GetPagerParams(defaultPageSize)
' Returns a Scripting.Dictionary with keys: page (int), pageSize (int)
Function GetPagerParams(Optional defaultPageSize)
    If IsMissing(defaultPageSize) Then defaultPageSize = 20
    Dim p, ps, allowed, found
    p = Request.QueryString("page")
    ps = Request.QueryString("pageSize")
    If Not IsNumeric(p) Then p = 0
    If Not IsNumeric(ps) Then ps = 0
    If CInt(p) < 1 Then p = 1
    If CInt(ps) < 1 Then ps = CInt(defaultPageSize)

    allowed = Array(10,20,50,100)
    found = False
    Dim i
    For i = 0 To UBound(allowed)
        If CInt(allowed(i)) = CInt(ps) Then found = True
    Next
    If Not found Then ps = CInt(defaultPageSize)

    Dim out
    Set out = CreateObject("Scripting.Dictionary")
    out.Add "page", CInt(p)
    out.Add "pageSize", CInt(ps)
    Set GetPagerParams = out
End Function

' Convenience: render pagination using current request params
Sub RenderPaginationFromRequest(total, Optional defaultPageSize, Optional options)
    If IsMissing(defaultPageSize) Then defaultPageSize = 20
    Dim pager
    Set pager = GetPagerParams(defaultPageSize)
    Call RenderPagination(total, pager.Item("page"), pager.Item("pageSize"), options)
End Sub
%>
