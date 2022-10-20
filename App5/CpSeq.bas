
Public Sub CpSeq(Optional control As IRibbonControl)
    If testing Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim currentCol As Integer
    currentCol = ActiveCell.Column


    'ActiveSheet.Rows(currentRow).Copy Destination:=ActiveSheet.Rows(currentRow + 1)

    ActiveSheet.Rows(currentRow).Select
    Selection.Copy
    ActiveSheet.Rows(currentRow + 1).Select
    Selection.Insert Shift:=xlDown
    Cells(currentRow + 1, currentCol).Select
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
    Application.ScreenUpdating = True
End Sub

