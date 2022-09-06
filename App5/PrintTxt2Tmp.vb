
Public Sub PrintTxt2Tmp(text As String, path As String)
    If testing Then
        Exit Sub
    End If

    Dim ff As Integer
    ff = FreeFile
    Open path For Append As #ff

    Print #ff, text
    Close #ff
End Sub

