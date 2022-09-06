
Public Sub ImpCsv()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    Dim filePath As String
    filePath = FnLstFil("C:\BAK\*.csv")


    Dim MyData As DataObject
    Set MyData = New DataObject
    MyData.SetText ReadLineByFile(filePath)
    MyData.PutInClipboard
    Set MyData = Nothing

    ActiveSheet.Paste

    Application.ScreenUpdating = True
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

