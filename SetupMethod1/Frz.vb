
Public Sub Frz()
    If testing Then Exit Sub
    Call UnHF
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 9
        .SplitRow = 1
        .FreezePanes = True
    End With

    'Range("A1").Select
    'Selection.AutoFilter
End Sub

