
Public Function CommonCheckSynWithM_OPEN()
    If testing Then Exit Function
    Dim closeFlag As Boolean
    closeFlag = False

    Dim activeName As String
    activeName = ActiveWorkbook.FullName
    
    If InStr(ActiveWorkbook.Sheets("Info").Range("A1"), "#") > 0 Then
        CommonCheckSynWithM_OPEN = closeFlag
        Exit Function
    End If
    
    If InStr(activeName, "http") = 1 Then
        CommonCheckSynWithM_OPEN = closeFlag
        Exit Function
    End If
    
    Dim fso As Object
    'Dim md As Object
    Dim fileObject, cFileObject, mFileObject As Object

    Dim cPath, mPath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(activeName)
    'Set md = fso.GetDrive(theDrive)
    
    Dim path As String
    Dim parameter As String
    'MsgBox Len(theDrive)
    'MsgBox theUser
    If Len(theDrive) >= 2 Then
    If InStr(activeName, "C:") > 0 Then
        mPath = Replace(activeName, "C:", theDrive)
        'If md.IsReady Then
        
        'If Dir(mPath) <> "" Then
        If fso.fileexists(mPath) Then
            Set mFileObject = fso.GetFile(mPath)
        
            If fileObject.DateLastModified < mFileObject.DateLastModified Then
                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
                    MyQuestionBox "You are not the author of this document and found another updated verion in share drive, Do U want to override from " & theDrive & " after close? Last modifier is " & GetWorkbookProperties(mPath, "Last Author"), "Yes", "No", 10
                    If confirmation = "Yes" Then
                        nexttime = Now() + TimeSerial(0, 0, 5)
                        Application.OnTime nexttime, "'CpFil2FilBk """ & mPath & """, """ & activeName & """, True'"

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

        'End If
    End If
    
    If InStr(activeName, theDrive) > 0 Then
        cPath = Replace(activeName, theDrive, "C:")

        'If Dir(cPath) <> "" Then
        If fso.fileexists(cPath) Then
            Set cFileObject = fso.GetFile(cPath)
        
            If fileObject.DateLastModified < cFileObject.DateLastModified Then
                If InStr(ActiveWorkbook.BuiltinDocumentProperties("Author").Value, theUser) = 0 Then
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
            MyQuestionBox "You didn't update a local copy of this document yet, Do U want to proceed now? ", "Yes", "No", 10
            If confirmation = "Yes" Then
                fso.copyfile activeName, cPath, True
            End If
        
            closeFlag = False
        End If
    
    End If
    End If
    Set fso = Nothing
    
    CommonCheckSynWithM_OPEN = closeFlag
End Function

