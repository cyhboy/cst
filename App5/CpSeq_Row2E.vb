
Public Sub CpSeq_Row2E()
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Call CpSeq
    Call Row2E

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

