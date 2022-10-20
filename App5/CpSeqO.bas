
Public Sub CpSeqO(Optional control As IRibbonControl)
    If testing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    Dim count As Integer, countNew As Integer
    'count = Range("A65536").End(xlUp).Row
    Dim endCount As Integer
    'endCount = 32767
    endCount = 3276

    For count = endCount To 1 Step -1
        If Cells(count, 1) <> "" And Cells(count, 1).EntireColumn.Hidden = False And Cells(count, 1).EntireRow.Hidden = False Then
            Exit For
        End If
    Next count

    For countNew = count To endCount
        If Cells(countNew, 1) = "" And Cells(countNew, 1).EntireColumn.Hidden = False And Cells(countNew, 1).EntireRow.Hidden = False Then
            Exit For
        End If
    Next countNew

    'If Cells(count, 1) <> vbNullString Then  'skip blank lines
    'ActiveSheet.Rows(count + 1).Insert
    ActiveSheet.Rows(count).Copy Destination:=ActiveSheet.Rows(countNew)
    'End If

    'count = Range("A65536").End(xlUp).Row
    Range("H" & countNew).Activate
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Cells(currentRow, 15) = "'" & LPad(Cells(currentRow - 1, 15) + 1, 4, "0")

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    Application.ScreenUpdating = True
End Sub

