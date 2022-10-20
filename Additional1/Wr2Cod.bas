
Public Sub Wr2Cod()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim resultStr As String
    Dim codeStr As String
    codeStr = Cells(currentRow, 10)

    If Trim(codeStr) = "" Then
        Exit Sub
    End If

    Dim comment As String
    comment = Cells(currentRow, 8)

    resultStr = comment & vbCrLf & codeStr

    Dim localFolder As String
    localFolder = Cells(currentRow, 9)

    If localFolder = "" Then
        localFolder = "C:\TMP\"
        Cells(currentRow, 9) = localFolder
    End If

    Call Fold

    Dim speicalFile As String
    speicalFile = Cells(currentRow, 11)

    If Trim(speicalFile) = "" Then
        speicalFile = "tmp_" & Format(Now, "yyyyMMdd-HHmmss") & ".sql"
        Cells(currentRow, 11) = speicalFile
    End If

    Dim filePath As String
    filePath = localFolder & speicalFile
    WriteTxt2Code resultStr, filePath
    MyMsgBox "coding finish", 3
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

End Sub

