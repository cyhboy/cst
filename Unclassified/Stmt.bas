
Public Sub Stmt()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim msg As String
    msg = Cells(currentRow, 10)
    
    MyMsgBox msg, 300
ErrorHandler:
        If Err.Number <> 0 Then
            MyMsgBox Err.Number & " " & Err.Description, 30
        End If
End Sub


