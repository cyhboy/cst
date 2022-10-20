
Public Sub MvFil2Fil(filePath1 As String, filePath2 As String, displayFlag As Boolean)
    If testing Then
        Exit Sub
    End If

    'On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.MoveFile filePath1, filePath2
    Set fso = Nothing
    If displayFlag Then
        MyMsgBox filePath1 & " to " & filePath2 & " moved", 5
    End If
    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub



