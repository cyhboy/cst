
Public Sub Caller2()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler
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
                RobotRunByParam "Caller2"
            End If
        Next curCell
        Exit Sub
    End If


    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim subNames As String
    subNames = Cells(currentRow, 8)

    Dim subNameArr As Variant
    subNameArr = Split(subNames, Chr(13) & Chr(10))

    Dim callerList As String
    callerList = ""


    Dim subName As String


    Dim funArr As Variant

    Dim i As Integer
    Dim j As Integer

    For i = 0 To UBound(subNameArr) - 1
        subName = Mid(subNameArr(i), InStrRev(subNameArr(i), "-") + 1)
        subName = Left(subName, InStr(subName, "{") - 1)

        funArr = CallerSignatures(subName)
        'MsgBox UBound(funArr)
        For j = 1 To UBound(funArr)
            If callerList = "" Then
                callerList = subNameArr(i) & "<-" & CStr(funArr(j))
            Else
                callerList = callerList & Chr(13) & Chr(10) & subNameArr(i) & "<-" & CStr(funArr(j))
            End If
        Next j

    Next i


    If callerList <> "" Then
        Cells(currentRow, 9) = callerList & Chr(13) & Chr(10)
    Else
        Cells(currentRow, 9) = ""
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

