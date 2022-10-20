
Public Sub Colr()
    If testing Then
        Exit Sub
    End If
    If ActiveSheet.Tab.ColorIndex >= 18 Or ActiveSheet.Tab.ColorIndex < 1 Then
        ActiveSheet.Tab.ColorIndex = 1
    Else
        ActiveSheet.Tab.ColorIndex = CInt(ActiveSheet.Tab.ColorIndex + 1)
    End If
End Sub

