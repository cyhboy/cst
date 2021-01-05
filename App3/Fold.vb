
Public Sub Fold()
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
                RobotRunByParam "Fold"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim strDirectory As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    strDirectory = Cells(currentRow, 9)
    CreateFolder strDirectory
    
'    Dim cell As Object
'    For Each cell In Selection.Cells
'        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
'            currentRow = cell.Row
'            strDirectory = Cells(currentRow, 9)
'            CreateFolder strDirectory
'        End If
'    Next
    
    'MsgBox "Create standard folder " & strDirectory & " successfully"
    'MsgBox "Create standard folder successfully"
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

