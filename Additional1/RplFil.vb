
Public Sub RplFil()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim n As Integer
    n = Selection.count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).count
    End If
    If n > 1 Then
        Dim curCell As Range
        For Each curCell In Selection
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                curCell.Select
                'MsgBox subName
                RobotRunByParam "RplFil"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim orgTxt As String
    Dim newTxt As String

    orgTxt = Cells(currentRow, 24)
    newTxt = Cells(currentRow, 25)

    If orgTxt = newTxt Then
        Exit Sub
    End If

    Dim localPath As String
    localPath = Cells(currentRow, 9)

    Dim fileName As String
    fileName = Cells(currentRow, 11)


    RplTxt4Fil localPath & fileName, orgTxt, newTxt

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

