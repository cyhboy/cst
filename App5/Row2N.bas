
Public Sub Row2N()
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim rng As Range

    Set rng = Selection

    rng.EntireRow.Select

    Selection.Cut

    rng.Offset(2, 0).EntireRow.Select

    Selection.Insert Shift:=xlDown

    ActiveCell.Offset(-1, 0).Select
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

