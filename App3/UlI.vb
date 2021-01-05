
Public Sub UlI()
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    Dim n As Integer
    n = Selection.Count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).Count
    End If
    If n > 1 Then
        Dim curCell As Range
        For Each curCell In Selection
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                curCell.Select
                'MsgBox subName
                RobotRunByParam "UlI"
            End If
        Next curCell
        Exit Sub
    End If
    
    Call FilDt
    UlParam False
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

