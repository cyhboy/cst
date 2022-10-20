
Public Sub PrintResult(text As String)
    If testing Then
        Exit Sub
    End If

    WriteTxt2Tmp text, "C:\BAK\interaction.log"
End Sub

