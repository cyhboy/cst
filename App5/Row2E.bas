
Public Sub Row2E(Optional control As IRibbonControl)
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim rng As Range
    Set rng = Selection
    Dim rw As Range

    Dim currentColumn As Integer
    currentColumn = ActiveCell.Column

    'For Each rw In rng.Rows

    Application.ScreenUpdating = False

    'Dim mycll As Range

    'Set mycll = Cells(rw.Row, 1)

    Dim count As Integer, countNew As Integer
    'Application.ScreenUpdating = False
    'count = Range("A1").End(xlDown).Row

    For count = 32767 To 1 Step -1
        If Cells(count, 1) <> "" And Cells(count, 1).EntireColumn.Hidden = False And Cells(count, 1).EntireRow.Hidden = False Then
            Exit For
        End If
    Next count

    'Exit Sub
    'MsgBox Len(Cells(count, 1))
    'MsgBox count

    For countNew = count To 32767
        If Cells(countNew, 1) = "" And Cells(countNew, 1).EntireColumn.Hidden = False And Cells(countNew, 1).EntireRow.Hidden = False Then
            Exit For
        End If
    Next countNew

    rng.EntireRow.Select

    Selection.Cut

    Cells(countNew, 1).EntireRow.Select

    Selection.Insert Shift:=xlDown

    ActiveCell.Offset(-1, currentColumn - 1).Select

    Application.ScreenUpdating = True

    'Next rw
    'Call InsSeq
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If

End Sub

