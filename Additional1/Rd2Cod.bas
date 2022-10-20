
Public Sub Rd2Cod()
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
                RobotRunByParam "Rd2Cod"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim comment As String
    comment = Cells(currentRow, 8)

    Dim localFolder As String
    localFolder = Cells(currentRow, 9)

    Dim fileName As String
    fileName = Cells(currentRow, 11)

    Dim fileContent As String
    fileContent = ReadLineByFile(localFolder & fileName)

    If InStr(fileContent, comment & vbCrLf) > 0 Then
        Cells(currentRow, 10) = Replace(fileContent, comment & vbCrLf, "")
    Else
        Cells(currentRow, 10) = fileContent
    End If

    Dim lineCnt As Integer
    lineCnt = CntSubstring(Cells(currentRow, 10), Chr(10))
    Cells(currentRow, 23) = lineCnt

ErrorHandler:
    If Err.Number <> 0 Then
        PrintResult "FAILED"
        MyMsgBox Err.Number & " " & Err.Description, 5
    Else
        If lineCnt > 2 Then
            PrintResult "NOT"
        Else
            PrintResult "END"
        End If
    End If
End Sub

