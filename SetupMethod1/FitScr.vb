
Public Sub FitScr(Optional control As IRibbonControl)
    If testing Then
        Exit Sub
    End If

    Application.WindowState = xlMaximized
    ActiveWindow.WindowState = xlMaximized
    Dim zoom As Double
    zoom = ActiveWindow.zoom

    Dim ww As Double
    Dim w As Double
    Dim cw As Double
    Dim x As Double
    'MsgBox ActiveWindow.Width
    'MsgBox ActiveWindow.UsableWidth
    ww = ActiveWindow.Width
    'ww = ActiveWindow.UsableWidth
    Dim sumDbl As Double
    sumDbl = 0
    Dim ratioAry As Variant
    ratioAry = Array(0.03, 0.04, 0.03, 0.03, 0.06, 0.065, 0.025, 0.075, 0.07, 0.12, 0.04, 0.055, 0.04, 0.04, 0.025, 0.03, 0.025, 0.025, 0.025, 0.025, 0.025, 0.025, 0.055)
    Dim i As Integer
    For i = 1 To 23
        sumDbl = sumDbl + ratioAry(i - 1)
        With Columns(i)
            w = .Width
            cw = .ColumnWidth
            x = ww * cw * 100 * ratioAry(i - 1) / w / zoom

            If x < 255 Then
                .ColumnWidth = x
            Else
                .ColumnWidth = 255
            End If

        End With
    Next i

    MyMsgBox "Total fill windows rate --> " & sumDbl, 5000

End Sub

