
Public Sub SynMZ()
    If testing Then
        Exit Sub
    End If

    Dim fso As Object
    Dim fileObject, cFileObject, mFileObject, obj As Object

    Dim cPath, mPath As String
    Dim iRet As Integer

    Dim activeName As String
    activeName = ActiveWorkbook.FullName

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    'MsgBox fso.Drives.count
    If fso.Drives.count > 1 Then
        For Each obj In fso.Drives
            On Error GoTo ErrorHandler
            'MsgBox obj.DriveType & obj.path
            'MsgBox theDrive
            If obj.DriveType = 3 Or obj.DriveType = 1 Then
                If obj.path = theDrive Or obj.path = "Z:" Then
                    If InStr(activeName, "C:") > 0 Then
                        mPath = Replace(activeName, "C:", obj.path)
                        'MsgBox mPath
                        If fso.fileexists(mPath) Then

                            Set mFileObject = fso.GetFile(mPath)

                            If fileObject.DateLastModified > mFileObject.DateLastModified Then
                                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                                    iRet = MsgBox("You are NOT the author of this document, Do U want to manually update " & obj.path & " as well? ", vbYesNo, "Question")
                                    If iRet = vbYes Then
                                        fso.copyfile activeName, mPath, True
                                    End If
                                End If
                            End If
                        Else
                            'MsgBox ActiveWorkbook.BuiltinDocumentProperties("Author").Value
                            If InStr(LCase(ActiveWorkbook.BuiltinDocumentProperties("Author").Value), theUser) > 0 Then
                                iRet = MsgBox("You are the author of this document, Do U want to manually append " & obj.path & " as well? ", vbYesNo, "Question")
                                If iRet = vbYes Then
                                    fso.copyfile activeName, mPath, True
                                End If
                            End If
                        End If
                    ElseIf InStr(activeName, obj.path) > 0 Then
                        cPath = Replace(activeName, obj.path, "C:")

                        If fso.fileexists(cPath) Then
                            Set cFileObject = fso.GetFile(cPath)

                            If fileObject.DateLastModified > cFileObject.DateLastModified Then

                                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
                                    iRet = MsgBox("You are the author of this document, Do U want to manually update C: as well? ", vbYesNo, "Question")
                                    If iRet = vbYes Then
                                        fso.copyfile activeName, cPath, True
                                    End If
                                End If
                            End If
                        Else
                            If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                                iRet = MsgBox("You are NOT the author of this document, Do U want to manually append C: as well? ", vbYesNo, "Question")
                                If iRet = vbYes Then
                                    fso.copyfile activeName, cPath, True
                                End If
                            End If
                        End If

                    End If
                End If
            End If

ErrorHandler:
            If Err.Number <> 0 Then
                MyMsgBox Err.Number & " " & Err.Description, 30
            End If
        Next obj
    Else
        '        If InStr(activeName, "C:") > 0 Then
        '            mPath = Replace(activeName, "C:", "\\10.15.76.73\common_oa")
        '            If fso.FileExists(mPath) Then
        '                Set mFileObject = fso.GetFile(mPath)
        '                If fileObject.DateLastModified > mFileObject.DateLastModified Then
        '                    If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
        '                        MyQuestionBox "You are NOT the author of this document, Do U want to manually update \\10.15.76.73\common_oa as well? ", "No", "Yes", 10
        '                        If confirmation = "Yes" Then
        '                            fso.copyfile activeName, mPath, True
        '                        End If
        '                    End If
        '                End If
        '            Else
        '                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) > 0 Then
        '                    MyQuestionBox "You are the author of this document, Do U want to manually append \\10.15.76.73\common_oa as well? ", "Yes", "No", 10
        '                    If confirmation = "Yes" Then
        '                        fso.copyfile activeName, mPath, True
        '                    End If
        '                End If
        '            End If
        '        End If
    End If
    Set fso = Nothing
End Sub

