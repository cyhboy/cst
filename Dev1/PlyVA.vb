
Public Sub PlyVA()
    If testing Then
        Exit Sub
    End If
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
                RobotRunByParam "PlyVA"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row


    'GoTo SECOND_STAGE
    Dim parameter As String
    parameter = Cells(currentRow, 10)
    Dim formatCode As String
    formatCode = CutStrByStartEnd(parameter, " best", "http", True)
    If InStr(parameter, "http") > 0 Then
        parameter = CutStrByStartEnd(parameter, "http", "$", True)
    Else
        parameter = ""
    End If

    Dim cmdStr As String
    ' cmdStr = "conda activate learn"
    ' cmdStr = cmdStr & " && " & "python C:\AppFiles\ipy\plyVA.py """ & parameter & """"
    ' after pyinstaller build the python file
    cmdStr = "C:\AppFiles\ipy\plyVA\plyVA.exe """ & parameter & """"
    Cells(currentRow, 18) = "'" & ShellRunResult(cmdStr, "C:\BAK\cmd.log", True)

    'SECOND_STAGE:

    Dim jsonStr As String
    jsonStr = Cells(currentRow, 18)
    jsonStr = CutStrByStartEnd(jsonStr, "{", "$", True)
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(jsonStr)
    Cells(currentRow, 1) = Json("subtitles")
    Cells(currentRow, 2) = Json("filesizeString")
    Cells(currentRow, 3) = Json("view_count")
    Cells(currentRow, 4) = "'" & Json("upload_date")
    Cells(currentRow, 8) = Replace(Cells(currentRow, 8), formatCode, " " & Json("formatCode") & " ")
    Cells(currentRow, 10) = Replace(Cells(currentRow, 10), formatCode, " " & Json("formatCode") & " ")
    Cells(currentRow, 13) = "'" & Json("videoFileName")

End Sub

