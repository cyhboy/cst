
Public Sub WriteTxt2Tmp(text As String, path As String)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Dim txt As Integer
    txt = FreeFile()
    Open path For Output As #txt
    Print #txt, text
    Close #txt
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

