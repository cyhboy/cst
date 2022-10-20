
Public Function GetResult()
    If testing Then
        Exit Function
    End If

    GetResult = Trim(ReadLineByFile("C:\BAK\interaction.log"))
End Function

