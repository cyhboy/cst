
Public Sub PrintMail(text As String)
    If testing Then
        Exit Sub
    End If

    WriteTxt2Tmp text, "C:\BAK\em.log"
End Sub


