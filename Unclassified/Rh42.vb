
Public Sub Rh42(Optional control As IRibbonControl)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
'    Dim n As Integer
'    n = Selection.count
'    If n > 1 Then
'        n = Selection.SpecialCells(xlCellTypeVisible).count
'    End If
'    If n > 1 Then
'        Dim curCell As Range
'        For Each curCell In Selection
'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
'                curCell.Select
'                'MsgBox subName
'                RobotRunByParam "Rh42"
'            End If
'        Next curCell
'        Exit Sub
'    End If
    
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.RowHeight = 42
    
    Call FitScr
    Call Frz
    Call RstCf
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

