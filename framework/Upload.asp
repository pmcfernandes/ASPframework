<!-- Upload.asp
   Simple Upload wrapper for Classic ASP
   Usage:
     Dim u: Set u = New Upload
     u.SaveRequestFiles Server.MapPath("/uploads")
     Dim files: Set files = u.Files
     For Each f In files
         Response.Write f.FileName & " saved to " & f.SavedPath & " (" & f.Size & " bytes)<br/>"
     Next
-->

<%
Class Upload
    Public Files ' array of dictionaries with keys: FileName, ContentType, Size, SavedPath, FieldName
    Public MaxFileSize ' bytes, 0 = unlimited
    Public AllowedTypes ' optional array of mime types allowed
    Public AllowedExtensions ' optional array of allowed file extensions (without dot), e.g. Array("jpg","png")
    Public DenyExtensions ' optional array of denied file extensions (without dot), e.g. Array("exe","asp")

    Private Sub Class_Initialize()
        ReDim Files(-1)
        MaxFileSize = 0
        Set AllowedTypes = Nothing
        Set AllowedExtensions = Nothing
        Set DenyExtensions = Nothing
    End Sub

    ' SanitizeFileName(baseName)
    ' Removes path segments and unsafe characters from a filename base.
    Private Function SanitizeFileName(s)
        On Error Resume Next
        If IsNull(s) Then
            SanitizeFileName = "file"
            Exit Function
        End If
        s = CStr(s)
        ' strip any path separators
        s = Replace(s, "/", "-")
        s = Replace(s, "\", "-")
        ' remove characters we don't want in filenames
        Dim i, ch, out
        out = ""
        For i = 1 To Len(s)
            ch = Mid(s, i, 1)
            Select Case Asc(ch)
                Case 48 To 57, 65 To 90, 97 To 122 ' 0-9, A-Z, a-z
                    out = out & ch
                Case 32 ' space -> hyphen
                    out = out & "-"
                Case 45, 95 ' - and _ allowed
                    out = out & ch
                Case Else
                    ' skip other chars
            End Select
        Next
        ' collapse multiple hyphens
        Do While InStr(out, "--") > 0
            out = Replace(out, "--", "-")
        Loop
        If Len(out) = 0 Then out = "file"
        SanitizeFileName = out
    End Function

    ' SaveRequestFiles(targetDir)
    ' Saves uploaded files from the current Request to targetDir. Returns number of saved files.
    Public Function SaveRequestFiles(targetDir)
        Dim savedCount
        savedCount = 0
        On Error Resume Next
        If Len(Trim(targetDir)) = 0 Then
            SaveRequestFiles = 0
            Exit Function
        End If
        If Right(targetDir,1) <> "\"" And Right(targetDir,1) <> "/" Then targetDir = targetDir & "\"
        Dim fso
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(targetDir) Then fso.CreateFolder(targetDir)

        Dim contentType
        contentType = Request.ServerVariables("CONTENT_TYPE")
        If InStr(LCase(contentType), "multipart/form-data") = 0 Then
            SaveRequestFiles = 0
            Exit Function
        End If

        ' Use common ADODB.Stream + BinaryRead parsing approach
        Dim raw
        Dim total
        total = Request.TotalBytes
        If total <= 0 Then
            SaveRequestFiles = 0
            Exit Function
        End If
        raw = Request.BinaryRead(total)

        ' Determine boundary
        Dim boundary
        boundary = "--" & Mid(contentType, InStr(contentType, "boundary=") + 9)
        If Len(boundary) = 2 Then
            SaveRequestFiles = 0
            Exit Function
        End If

        Dim arrParts
        arrParts = Split(raw, boundary)
        Dim i
        For i = 0 To UBound(arrParts)
            Dim part
            part = arrParts(i)
            If Len(Trim(part)) = 0 Or InStr(part, "--") = 1 Then Continue For
            ' separate headers and body
            Dim headerEnd
            headerEnd = InStr(part, vbCrLf & vbCrLf)
            If headerEnd = 0 Then Continue For
            Dim headersText, bodyBin
            headersText = Left(part, headerEnd-1)
            bodyBin = Mid(part, headerEnd+4)
            ' parse Content-Disposition
            Dim dispLine
            dispLine = ""
            If InStr(LCase(headersText), "content-disposition:") > 0 Then
                Dim lines
                lines = Split(headersText, vbCrLf)
                Dim ln
                For Each ln In lines
                    If InStr(LCase(ln), "content-disposition:") > 0 Then
                        dispLine = Trim(ln)
                        Exit For
                    End If
                Next
            End If
            If dispLine = "" Then Continue For
            ' extract filename and fieldname
            Dim fname, fieldname
            fname = ""
            fieldname = ""
            If InStr(LCase(dispLine), "filename=") > 0 Then
                fname = Mid(dispLine, InStr(LCase(dispLine), "filename=")+9)
                fname = Replace(fname, '"', "")
            End If
            If InStr(LCase(dispLine), "name=") > 0 Then
                Dim tmp
                tmp = Mid(dispLine, InStr(LCase(dispLine), "name=")+5)
                tmp = Replace(tmp, '"', "")
                fieldname = tmp
            End If

            ' parse content-type header
            Dim pctype
            pctype = "application/octet-stream"
            If InStr(LCase(headersText), "content-type:") > 0 Then
                Dim lns
                lns = Split(headersText, vbCrLf)
                Dim l
                For Each l In lns
                    If InStr(LCase(l), "content-type:") > 0 Then
                        pctype = Trim(Mid(l, InStr(LCase(l), "content-type:")+13))
                        Exit For
                    End If
                Next
            End If

            If fname <> "" Then
                ' validate size heuristic
                Dim sz
                sz = LenB(bodyBin)
                If MaxFileSize > 0 And sz > CLng(MaxFileSize) Then
                    ' skip
                Else
                            Dim extNoDot
                            extNoDot = LCase(fso.GetExtensionName(fileOnly))

                            ' extension allow/deny checks
                            If Not AllowedExtensions Is Nothing Then
                                Dim allowedExtOK
                                allowedExtOK = False
                                Dim aext
                                For Each aext In AllowedExtensions
                                    If LCase(aext) = extNoDot Then allowedExtOK = True
                                Next
                                If Not allowedExtOK Then
                                    ' skip due to extension not in allow list
                                    GoTo NextPart_Skip_Save
                                End If
                            ElseIf Not DenyExtensions Is Nothing Then
                                Dim denyExtFound
                                denyExtFound = False
                                Dim dext
                                For Each dext In DenyExtensions
                                    If LCase(dext) = extNoDot Then denyExtFound = True
                                Next
                                If denyExtFound Then
                                    ' skip due to deny list
                                    GoTo NextPart_Skip_Save
                                End If
                            End If

                            If Not AllowedTypes Is Nothing Then
                        Dim ok
                        ok = False
                        Dim t
                        For Each t In AllowedTypes
                            If LCase(t) = LCase(pctype) Then ok = True
                        Next
                        If Not ok Then
                            ' skip
                        Else
                            ' save file
                            Dim saveName, fullPath
                            ' sanitize original filename and build a unique saved name
                            Dim origFileName, fileOnly, baseName, ext, safeBase, uniqueSuffix
                            origFileName = fname
                            fileOnly = fso.GetFileName(origFileName)
                            baseName = fso.GetBaseName(fileOnly)
                            ext = fso.GetExtensionName(fileOnly)
                            If ext <> "" Then ext = "." & ext
                            safeBase = SanitizeFileName(baseName)
                            Randomize
                            uniqueSuffix = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2) & "_" & _
                                           Right("0" & Hour(Now()),2) & Right("0" & Minute(Now()),2) & Right("0" & Second(Now()),2) & "_" & CInt(Rnd()*1000000)
                            saveName = safeBase & "_" & uniqueSuffix & ext
                            fullPath = targetDir & saveName
                            Dim stm
                            Set stm = Server.CreateObject("ADODB.Stream")
                            stm.Type = 1 ' adTypeBinary
                            stm.Open
                            stm.Write bodyBin
                            stm.SaveToFile fullPath, 2 ' adSaveCreateOverWrite
                            stm.Close
                            Set stm = Nothing

                            ' record file meta
                            Dim d
                            Set d = CreateObject("Scripting.Dictionary")
                            d.Add "FileName", fname ' original filename provided by client
                            d.Add "SavedName", saveName ' sanitized, unique filename on disk
                            d.Add "SavedPath", fullPath
                            d.Add "Size", sz
                            d.Add "ContentType", pctype
                            d.Add "FieldName", fieldname
                            Dim idx
                            idx = UBound(Files) + 1
                            ReDim Preserve Files(idx)
                            Set Files(idx) = d
                            savedCount = savedCount + 1
                            End If
NextPart_Skip_Save:
                    Else
                        ' save without mime check
                        Dim saveName2, fullPath2
                        Dim extNoDot2
                        extNoDot2 = LCase(fso.GetExtensionName(fileOnly2))

                        ' extension allow/deny checks for branch without mime check
                        If Not AllowedExtensions Is Nothing Then
                            Dim allowedExtOK2
                            allowedExtOK2 = False
                            Dim aext2
                            For Each aext2 In AllowedExtensions
                                If LCase(aext2) = extNoDot2 Then allowedExtOK2 = True
                            Next
                            If Not allowedExtOK2 Then
                                GoTo NextPart2_Skip_Save
                            End If
                        ElseIf Not DenyExtensions Is Nothing Then
                            Dim denyExtFound2
                            denyExtFound2 = False
                            Dim dext2
                            For Each dext2 In DenyExtensions
                                If LCase(dext2) = extNoDot2 Then denyExtFound2 = True
                            Next
                            If denyExtFound2 Then
                                GoTo NextPart2_Skip_Save
                            End If
                        End If
                        ' sanitize original filename and build a unique saved name
                        Dim origFileName2, fileOnly2, baseName2, ext2, safeBase2, uniqueSuffix2
                        origFileName2 = fname
                        fileOnly2 = fso.GetFileName(origFileName2)
                        baseName2 = fso.GetBaseName(fileOnly2)
                        ext2 = fso.GetExtensionName(fileOnly2)
                        If ext2 <> "" Then ext2 = "." & ext2
                        safeBase2 = SanitizeFileName(baseName2)
                        Randomize
                        uniqueSuffix2 = Year(Now()) & Right("0" & Month(Now()),2) & Right("0" & Day(Now()),2) & "_" & _
                                        Right("0" & Hour(Now()),2) & Right("0" & Minute(Now()),2) & Right("0" & Second(Now()),2) & "_" & CInt(Rnd()*1000000)
                        saveName2 = safeBase2 & "_" & uniqueSuffix2 & ext2
                        fullPath2 = targetDir & saveName2
                        Dim stm2
                        Set stm2 = Server.CreateObject("ADODB.Stream")
                        stm2.Type = 1 ' adTypeBinary
                        stm2.Open
                        stm2.Write bodyBin
                        stm2.SaveToFile fullPath2, 2 ' adSaveCreateOverWrite
                        stm2.Close
                        Set stm2 = Nothing

                        Dim d2
                        Set d2 = CreateObject("Scripting.Dictionary")
                        d2.Add "FileName", fname ' original filename
                        d2.Add "SavedName", saveName2 ' sanitized unique filename
                        d2.Add "SavedPath", fullPath2
                        d2.Add "Size", LenB(bodyBin)
                        d2.Add "ContentType", pctype
                        d2.Add "FieldName", fieldname
                        Dim idx2
                        idx2 = UBound(Files) + 1
                        ReDim Preserve Files(idx2)
                        Set Files(idx2) = d2
                        savedCount = savedCount + 1
NextPart2_Skip_Save:
                    End If
                End If
            End If
        Next

        SaveRequestFiles = savedCount
    End Function

    ' Clear stored files metadata
    Public Sub Clear()
        ReDim Files(-1)
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        Dim i
        For i = 0 To UBound(Files)
            Set Files(i) = Nothing
        Next
    End Sub
End Class
%>
