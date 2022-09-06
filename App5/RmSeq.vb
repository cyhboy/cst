
Public Sub RmSeq()
    If testing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim count As Integer, countNew As Integer
    'Application.ScreenUpdating = False
    'count = Range("A1").End(xlDown).Row

    For count = 32767 To 1 Step -1
        If Cells(count, 1) <> "" And Cells(count, 1).EntireColumn.Hidden = False And Cells(count, 1).EntireRow.Hidden = False Then
            Exit For
        End If
    Next count

    Rows(count).Delete Shift:=xlUp

    ActiveCell.Offset(-1, 0).Select

    'Call InsSeq
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    Application.ScreenUpdating = True
End Sub

