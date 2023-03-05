
Private Sub Workbook_Open()
    'On Error GoTo ErrorHandler

    skipping = False
    
    'Exit Sub
    Dim fso As Object
    Dim cFileObject As Object
    Dim mFileObject As Object
    Dim obj As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim updateMacroPath As String
    Dim updateUiPath As String
    Dim updateUiPathVer As String
    
    Dim theFolder As String
    
    Dim scriptPath As String
    Dim scriptParameter As String
    
    Dim finalUiPath As String
    finalUiPath = Environ("USERPROFILE") & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    
    Dim finalMacroPath As String
    finalMacroPath = ThisWorkbook.FullName
    
    theDrive = CommonGetTheDrive()
    'MsgBox theDrive
    'theUser = RespExtMail(Environ$("username"), "EXTERNAL_MAIL")
    theUser = Environ$("username")
    
    'theUser = extMail(Environ$("username"))
    
    'MsgBox theUser
    'MsgBox ReadEnv("%PROGRAMFILES%")
    'MsgBox Environ("AppData")
    'MsgBox Environ("USERPROFILE")
    'MsgBox ThisWorkbook.FullName
    Dim copyMacroPathVer As String
    copyMacroPathVer = "C:\AppFiles\cst.xlam"
    Set cFileObject = fso.GetFile(copyMacroPathVer)

    Dim mFileDate As Date
    Dim cFileDate As Date

    mFileDate = DateAdd("yyyy", -5, Now)
    cFileDate = DateAdd("yyyy", -5, Now)
    
    cFileDate = cFileObject.DateLastModified

    For Each obj In fso.Drives()
        'MsgBox obj.path & " " & obj.DriveType
        'If obj.DriveType = 3 Then
        If obj.path <> "C:" Then
            'If Dir(obj.path & "\AppFiles\SupportSetup\cst.xlam") <> "" Then
            If fso.fileexists(obj.path & "\AppFiles\SupportSetup\cst.xlam") Then
            
                Set mFileObject = fso.GetFile(obj.path & "\AppFiles\SupportSetup\cst.xlam")
                If mFileObject.DateLastModified > mFileDate Then
                    mFileDate = mFileObject.DateLastModified
                    theFolder = mFileObject.parentFolder
                    updateMacroPath = obj.path & "\AppFiles\SupportSetup\cst.xlam"
                    updateUiPath = obj.path & "\AppFiles\SupportSetup\Excel.officeUI"
                    updateUiPathVer = obj.path & "\AppFiles\SupportSetup\" & "Excel_" & Environ$("username") & ".officeUI"
                End If
            End If
        End If
    Next obj

    Set mFileObject = Nothing
    Set cFileObject = Nothing
    Set fso = Nothing

    If (mFileDate - cFileDate > 0) Then
        MyMsgBox "Dear CST Users, Thanks for choosing common support toolkits for your daily work. You now was recommended to upgrade to a new CST version, Please free 1 min to close your office suites and double click " & theFolder & "\install.bat. Thanks very much in deep. ", 10

        'CpFil2Fil updateUiPath, finalUiPath, False
        'CpFil2Fil finalUiPath, updateUiPathVer, False
        'CpFil2Fil updateMacroPath, copyMacroPathVer, False

        'scriptPath = "WScript.exe C:\AppFiles\WaitThenRunHiddenJob.vbs "
        'scriptParameter = """cmd.exe /C copy /Y %22" & updateMacroPath & "%22" & " " & "%22" & finalMacroPath & "%22""" & " " & """5000"""
        'ShellRunHide scriptPath & scriptParameter
        'Sleep 1000
        'ThisWorkbook.Saved = True
        'ThisWorkbook.Close
    End If
    
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox "Dear CST Users, When you see this message, The initialization of Common Support Toolkits may encounter some abnormal, It would not affect your daily excel operation, Be patience and try to dump this screen to CST Support, Thanks much. " & Err.Number & " " & Err.Description, 15
    End If

End Sub


