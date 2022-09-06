
Public Sub CpSeq_Row2N()
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Call CpSeq
    Call Row2N

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

