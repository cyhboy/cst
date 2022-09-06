
Public Sub PrintLog(text As String)
    If testing Then
        Exit Sub
    End If

    Dim ff As Integer
    ff = FreeFile
    Open "C:\BAK\cst.log" For Append As #ff

    Print #ff, text
    Close #ff
End Sub

