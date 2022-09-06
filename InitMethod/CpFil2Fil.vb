
Public Sub CpFil2Fil(filePath1 As String, filePath2 As String, displayFlag As Boolean, overrideFlag As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.fileexists(filePath2) Or overrideFlag Then
        fso.copyfile filePath1, filePath2
    End If

    Set fso = Nothing
    If displayFlag Then
        MyMsgBox filePath1 & " to " & filePath2 & " copied", 5
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

