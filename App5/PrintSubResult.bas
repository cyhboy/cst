
Public Sub PrintSubResult(text As String)
    If testing Then
        Exit Sub
    End If

    WriteTxt2Tmp text, "C:\BAK\subinteraction.log"
End Sub

