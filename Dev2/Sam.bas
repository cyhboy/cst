
Public Sub Sam()
    If testing Then
        Exit Sub
    End If
    Dim aSize As Variant
    Dim totalSize As Double
    
    Dim curCell As Range
    For Each curCell In Selection
        If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
            aSize = curCell.Value
            If IsNumeric(aSize) Then
                aSize = CDbl(aSize)
            Else
                aSize = SearchRegxKwInStr(CStr(aSize), "([\+\-]*[0-9]+\.*[0-9]*)")
                aSize = CDbl(aSize)
            End If
            totalSize = totalSize + aSize
        End If
    Next curCell
    
    MsgBox totalSize
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 20
    End If
End Sub

