
Public Sub UnHF()
    ' unhide and unfilter
    If testing Then
        Exit Sub
    End If
    Dim currentws As Worksheet
    Set currentws = ActiveWorkbook.ActiveSheet

    currentws.Cells.Select
    Selection.EntireColumn.Hidden = False
    Selection.EntireRow.Hidden = False

    If currentws.AutoFilterMode = True Then
        currentws.Rows("1:1").Select
        currentws.AutoFilterMode = False
    End If

    currentws.Range("A1").Select
    If currentws.Cells(1, 1) <> "" Then
        Selection.AutoFilter
    End If
End Sub

