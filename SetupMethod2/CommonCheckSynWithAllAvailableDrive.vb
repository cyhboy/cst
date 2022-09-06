
Public Sub CommonCheckSynWithAllAvailableDrive()
    If testing Then
        Exit Sub
    End If

    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then Exit Sub

        If skipping Then
            Exit Sub
        End If

        Dim fso As Object
        Dim fileObject, cFileObject, mFileObject, obj As Object
        Dim cPath, mPath As String
        Dim idleCnt As Long
        Dim idleFlag As Boolean

        If Application.Version = "12.0" Then
            idleCnt = 7
            idleFlag = False
        Else
            idleFlag = True
            idleCnt = 7
        End If

        Dim activeName As String
        activeName = ActiveWorkbook.FullName

        Set fso = CreateObject("Scripting.FileSystemObject")
        Set fileObject = fso.GetFile(activeName)

        If fso.Drives.count >= 2 Then
            For Each obj In fso.Drives
                On Error GoTo ErrorHandler
                'MsgBox obj.path
                'If obj.DriveType = 3 Then
                If obj.path <> "C:" Then

                    If InStr(activeName, "C:") > 0 Then
                        mPath = Replace(activeName, "C:", obj.path)
                        'If Dir(mPath) <> "" Then
                        If fso.fileexists(mPath) Then
                            Set mFileObject = fso.GetFile(mPath)
                            If fileObject.DateLastModified > mFileObject.DateLastModified Then

                                If InStr(LCase(ActiveWorkbook.BuiltinDocumentProperties("Author").Value), theUser) > 0 Then
                                    If idleFlag Then
                                        MyQuestionBox "You are the author of this document, Do U want to update " & obj.path & " as well? ", "Yes", "No", idleCnt
                                        If confirmation = "Yes" Then
                                            fso.copyfile activeName, mPath, True
                                        End If
                                    Else
                                        fso.copyfile activeName, mPath, True
                                    End If
                                End If
                            End If

                        End If
                    ElseIf InStr(activeName, obj.path) > 0 Then
                        cPath = Replace(activeName, obj.path, "C:")
                        'If Dir(cPath) <> "" Then
                        If fso.fileexists(cPath) Then
                            Set cFileObject = fso.GetFile(cPath)
                            If fileObject.DateLastModified > cFileObject.DateLastModified Then
                                'If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) >= 0 Then
                                If idleFlag Then
                                    MyQuestionBox "You are not the author of this document, Do U want to update C: as well? ", "Yes", "No", idleCnt
                                    If confirmation = "Yes" Then
                                        fso.copyfile activeName, cPath, True
                                    End If
                                Else
                                    fso.copyfile activeName, cPath, True
                                End If
                                'End If
                            End If
                        End If
                    End If
                End If
ErrorHandler:
                If Err.Number <> 0 Then
                    MyMsgBox Err.Number & " " & Err.Description, 7
                End If
            Next obj
        Else
            If Len(theDrive) = 2 Then
                If InStr(activeName, "C:") > 0 Then
                    mPath = Replace(activeName, "C:", theDrive)
                    'If Dir(mPath) <> "" Then
                    If fso.fileexists(mPath) Then
                        Set mFileObject = fso.GetFile(mPath)
                        If fileObject.DateLastModified > mFileObject.DateLastModified Then
                            If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                                If idleFlag Then
                                    MyQuestionBox "You are the author of this document, Do U want to update " & theDrive & " as well? ", "Yes", "No", idleCnt
                                    If confirmation = "Yes" Then
                                        fso.copyfile activeName, mPath, True
                                    End If
                                Else
                                    fso.copyfile activeName, mPath, True
                                End If
                            End If
                        End If

                    End If
                End If
            End If
        End If

        Set fso = Nothing
End Sub

