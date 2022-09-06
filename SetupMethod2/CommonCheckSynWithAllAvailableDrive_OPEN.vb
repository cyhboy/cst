
Public Function CommonCheckSynWithAllAvailableDrive_OPEN()
    If testing Then
        Exit Function
    End If

    Dim closeFlag As Boolean
    closeFlag = False

    Dim activeName As String
    activeName = ActiveWorkbook.FullName

    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then
        CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
        Exit Function
    End If

    If InStr(activeName, "http") = 1 Then
        CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
        Exit Function
    End If

    Dim fso As Object
    Dim fileObject, cFileObject, mFileObject As Object

    Dim cPath, mPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)

    Dim path As String
    Dim parameter As String

    Dim obj As Object
    For Each obj In fso.Drives
        If obj.DriveType = 3 Or obj.DriveType = 1 Then
            If InStr(activeName, "C:") > 0 Then
                mPath = Replace(activeName, "C:", obj.path)

                'If Dir(mPath) <> "" Then
                If fso.fileexists(mPath) Then
                    Set mFileObject = fso.GetFile(mPath)
                    MsgBox fileObject.DateLastModified
                    MsgBox mFileObject.DateLastModified
                    If fileObject.DateLastModified < mFileObject.DateLastModified Then
                        If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                            MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & obj.path & " after close? Last modifier is " & GetWorkbookProperties(mPath, "Last Author"), "Yes", "No", 10
                            If confirmation = "Yes" Then
                                nexttime = Now() + TimeSerial(0, 0, 5)
                                Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"
                                closeFlag = True
                                Exit For
                            Else
                                closeFlag = False
                                Exit For
                            End If
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If
            ElseIf InStr(activeName, obj.path) > 0 Then
                cPath = Replace(activeName, obj.path, "C:")

                'If Dir(cPath) <> "" Then
                If fso.fileexists(cPath) Then
                    Set cFileObject = fso.GetFile(cPath)

                    If fileObject.DateLastModified < cFileObject.DateLastModified Then
                        If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) < 0 Then
                            MyQuestionBox "You are not the author of this document and found another updated verion in C drive, Do U want to override from C: after close? Last modifier is " & GetWorkbookProperties(cPath, "Last Author"), "Yes", "No", 10
                            If confirmation = "Yes" Then
                                nexttime = Now() + TimeSerial(0, 0, 5)
                                Application.OnTime nexttime, "'CpFil2FilBk """ & cPath & """, """ & activeName & """, True'"
                                closeFlag = True
                            Else
                                closeFlag = False
                            End If
                        Else
                            closeFlag = False
                        End If
                    Else
                        closeFlag = False
                    End If
                Else
                    closeFlag = False
                End If

            End If
        End If

    Next obj

    Set fso = Nothing
    CommonCheckSynWithAllAvailableDrive_OPEN = closeFlag
End Function


