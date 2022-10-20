
Public Function IsBackgroudColor(colorValue As Long) As Boolean
    If testing Then
        Exit Function
    End If
    Dim redVal As Integer
    Dim greenVal As Integer
    Dim blueVal As Integer
    redVal = colorValue Mod 256
    greenVal = (colorValue \ 256) Mod 256
    blueVal = colorValue \ 65536
    If redVal + greenVal + blueVal >= 255 * 3 * 0.8 Then
        IsBackgroudColor = True
    Else
        IsBackgroudColor = False
    End If
End Function


