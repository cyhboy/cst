
Public Sub LstFil()
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
                RobotRunByParam "LstFil"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim localFolder As String
    Dim keyword As String
    localFolder = Cells(currentRow, 9)
    keyword = Cells(currentRow, 13)
    
    If EndsWith(keyword, ".mp4") Then
        keyword = Left(keyword, Len(keyword) - 3)
    ElseIf EndsWith(keyword, ".webm") Then
        keyword = Left(keyword, Len(keyword) - 4)
    End If
    

    
    Dim act As String
    act = Cells(currentRow, 16)

    '    If EndsWith(localFolder, "\") Then
    '        localFolder = Left(localFolder, Len(localFolder) - 1)
    '    End If

    ' MsgBox localFolder & keyword

    Dim currFilename As String
    currFilename = Cells(currentRow, 11)

    Dim lstFilename As String
    
    Dim date0 As Date
    Dim date1 As Date
    date0 = DateAdd("yyyy", -20, Now)
    date1 = date0
    
    Dim fileList As Variant
    ' MsgBox keyword
    ' fileList = GetFileList(localFolder & keyword & "*")
    fileList = GetFileList(localFolder & "*" & keyword & "*")
    ' MsgBox TypeName(fileList)
    ' MsgBox VarType(fileList)
    ' MsgBox Len(fileList)
    ' Exit Sub

    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myFileObj As Object
    Dim myFile As Variant
    For Each myFile In fileList
        'If CStr(myFile) <> currFilename Then
        Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
        If myFileObj.DateLastModified > date1 And Not InStr(myFileObj.Name, ".temp") > 0 And Not InStr(myFileObj.Name, ".f") > 0 And Not EndsWith(myFileObj.Name, ".srt") And Not EndsWith(myFileObj.Name, ".part") And Not EndsWith(myFileObj.Name, ".ytdl") Then
            date1 = myFileObj.DateLastModified
            'If rtnType = 1 Then filename1 = myFile.path
            'If rtnType = 2 Then filename1 = myFile.Name
            lstFilename = myFileObj.Name
        End If
        'End If
    Next myFile
    Set fso = Nothing
    If lstFilename <> "" Then
        Cells(currentRow, 11) = "'" & lstFilename
    End If
    If InStr(act, "EmRe") = 0 Then
        Cells(currentRow, 12) = date1
    End If
    

ErrorHandler:
    If Err.Number <> 0 Then
        ' MyMsgBox Err.Number & " " & Err.Description, 5
        If Cells(currentRow, 11) = "" Then
            Cells(currentRow, 11) = Err.Description
        End If

        If DateDiff("s", date0, date1) = 0 And Not EndsWith(keyword, "].") Then

            Dim findStr As String
            Dim rplStr As String
            findStr = StrReverse(CutStrByStartEnd(StrReverse(keyword), ".", "-", False, True))
            If Len(findStr) = 11 Then
                findStr = "-" & findStr
            End If
            If Len(findStr) = 12 Then
                rplStr = " [" & Right(findStr, 11) & "]"
                Cells(currentRow, 13) = Replace(Cells(currentRow, 13), findStr, rplStr)
                Call LstFil
            End If
        End If

    End If
End Sub

