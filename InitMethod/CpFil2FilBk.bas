
Public Sub CpFil2FilBk(filePath1 As String, filePath2 As String, displayFlag As Boolean)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.fileexists(filePath1) Then
        MyMsgBox filePath1 & " (failed) to " & filePath2, 5
    Else
        If Not fso.fileexists(filePath2) Then
            MyMsgBox filePath1 & " to " & filePath2 & " (failed)", 5
        End If
    End If

    Dim result As Variant
    result = fso.copyfile(filePath1, filePath2)
    Set fso = Nothing

    If displayFlag Then
        If result = "" Then
            MyMsgBox filePath1 & " to " & filePath2 & " copied", 5
        End If
    End If
    Application.Workbooks.Open filePath2

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

