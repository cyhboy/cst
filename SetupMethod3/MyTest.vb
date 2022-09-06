
Public Sub MyTest()
    If testing Then
        Exit Sub
    End If
    
    'MsgBox MatchRegx("ListNodes = ListNodes(node2, procName, iCol + 1, display)", "[=|>|<|&|-|,|+|f] ([A-Z][^ '\.,#\$\(=\\]+)\([^ ,]*(, [^ ,]+)*\)")
    'MsgBox MatchRegx("ListNodes = ListNodes(node2, procName, iCol + 1, display)", "[=|>|<|&|-|,|+|f] ([A-Z][^ '\.,#\$\(=\\]+)\([^,]*(, [^,]+)*\)")
    'MsgBox CntExeRunning("node.exe")

    'ActiveSheet.AutoFilter.ApplyFilter
    'Dim vCell As Range
    'Set vCell = ActiveSheet.AutoFilter.Range.offset(1, 0).SpecialCells(xlCellTypeVisible).Cells(1, 4)
    'vCell.Select

    'Dim currentRow As Integer
    'currentRow = ActiveCell.row

    'Application.ScreenUpdating = False

    'Cells(currentRow, 10) = "TEST"

    'MsgBox Cells(currentRow, 10) = Cells(currentRow, 11)

    'Application.ScreenUpdating = True

    'Dim tmpStr As String
    'tmpStr = "<span style='color:black'>01143211101<o:p></o:p><span style='color:black'>02663762301<o:p></o:p></span>"
    'MsgBox SearchRegxKwInStr(tmpStr, "([^0-9|,|\.|'])(0[0-9])")
    'MsgBox RplRegx(tmpStr, "([^0-9|,|\.|'])(0[0-9])", "$1'$2")

    'MyMsgBox Environ("username"), 5

    'Sleep 5000

    'MyQuestionBox Environ("username"), "Yes", "No", 5

    'MsgBox confirmation

    'MsgBox GetTickCount

    'Sleep 5000

    'MsgBox Environ("username")
    'MsgBox Cells(currentRow, 1)
    'MsgBox "A" > "*"

    'PrintLog (RemoveBlankLine(HtmlToText(GetHtmlByIe("http://aiahk-jira.aia.biz/browse/AIAPT-1173"))) & Chr(10))

    'PrintLog (GetTxtByIe("http://aiahk-jira.aia.biz/browse/AIAPT-1173"))
    'Exit Sub

    'MsgBox ActiveCell.Address
    'maxCall = maxCall - 1
    'MsgBox maxCall
    'Exit Sub

    'MsgBox Len(vbCrLf)
    'MsgBox Weekday(Date)

    'MsgBox Date

    'Rows("2:2").Select
    'Range(Selection, Selection.End(xlDown)).Select
    'Selection.RowHeight = 42

    'Dim filePath As String
    'filePath = Cells(currentRow, 9) & Cells(currentRow, 11)
    'MsgBox LastModDate("C:\BAK")

End Sub


