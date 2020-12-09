
Public Sub Robot(Optional control As IRibbonControl)
    If testing Then Exit Sub
    
    'On Error GoTo ErrorHandler
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
                RobotRunByParam "Robot"
            End If
        Next curCell
        Exit Sub
    End If

    Dim mycll As Excel.Range
    
    Set mycll = ActiveCell
    
    Dim comms As String
    'Dim interval As Long

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    comms = Cells(currentRow, 16)

    Dim arr
    arr = Split(comms, " ")
    
    Dim i As Integer
    
    Dim tmp_comm As String
    Dim comm As String
    
    For i = 0 To UBound(arr)
        tmp_comm = comm
        comm = arr(i)
        If InStr(comm, "#") = 0 Then
            If InStr(comms, "CpSeqO") = 0 And InStr(comms, "CpSeq") = 0 Then
                mycll.Select
            End If
            RobotRunByParam comm
        End If
    Next

    Exit Sub
    
'ErrorHandler:
'    If Err.Number <> 0 Then
'       MyMsgBox Err.Number & " " & Err.Description, 30
'    End If
End Sub

