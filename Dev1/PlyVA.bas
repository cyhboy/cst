
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
    
    Dim resultStr As String
    resultStr = Cells(currentRow, 10)
    
    Dim formatCode As String
    formatCode = CutStrByStartEnd(parameter, " best", "http", True)
    
    Dim origFile As String
    origFile = CutStrByStartEnd(parameter, "::ffmpeg -i """, """")

    Dim audioFile As String
    audioFile = CutStrByStartEnd(parameter, " -acodec copy """, """")
    
    If InStr(parameter, "http") > 0 Then
        parameter = CutStrByStartEnd(parameter, "http", "$$", True)
        If InStr(parameter, vbCrLf) > 0 Then
            parameter = CutStrByStartEnd(parameter, "http", vbCrLf, True)
        End If
    Else
        parameter = ""
    End If
    
    'MsgBox parameter
    'Exit Sub
    
    Dim cmdStr As String
    ' cmdStr = "conda activate learn"
    ' cmdStr = cmdStr & " && " & "python C:\AppFiles\ipy\plyVA.py """ & parameter & """"
    ' after pyinstaller build the python file
    cmdStr = "C:\AppFiles\ipy\plyVA\plyVA.exe """ & parameter & """"
    Cells(currentRow, 18) = "'" & ShellRunResult(cmdStr, "C:\BAK\cmd.log", True)

    'SECOND_STAGE:

    Dim jsonStr As String
    jsonStr = Cells(currentRow, 18)
    jsonStr = CutStrByStartEnd(jsonStr, "{", "$$", True)
    Dim Json As Object
    Set Json = JsonConverter.ParseJson(jsonStr)
    Dim cmdResult As String
    Dim subtitles As String
    subtitles = Json("subtitles")
    
    Cells(currentRow, 1) = subtitles
    Cells(currentRow, 2) = Json("filesizeString")
    Cells(currentRow, 3) = Json("view_count")
    Cells(currentRow, 4) = "'" & Json("upload_date")
    
    Dim rplFormatCode As String
    rplFormatCode = " " & Json("formatCode") & " "
    
    If Not (subtitles = "subtitles0[]" Or subtitles = "subtitles0" Or subtitles = "subtitlesErr" Or subtitles = "subtitlesNil") Then
        rplFormatCode = rplFormatCode & "--write-sub --sub-lang en,en-US,en-GB,zh,zh-CN,zh-HK,zh-TW,zh-Hans,zh-Hant --convert-subs srt "
    End If
    
    resultStr = Replace(resultStr, formatCode, rplFormatCode)
    resultStr = Replace(resultStr, origFile, Json("videoFileName"))
    
    Dim extStr As String
    extStr = Json("videoFileName")
    extStr = Right(extStr, Len(extStr) - InStrRev(extStr, ".") + 1)
    resultStr = Replace(resultStr, audioFile, Replace(Json("videoFileName"), extStr, ".opus"))
    Cells(currentRow, 10) = resultStr
    Cells(currentRow, 8) = resultStr
    
    Cells(currentRow, 13) = "'" & Json("videoFileName")

End Sub

