
Public Sub Wr2Fil()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim resultStr As String
    resultStr = Cells(currentRow, 10)

    Dim localFolder As String
    localFolder = Cells(currentRow, 9)

    If localFolder = "" Then
        localFolder = "C:\TMP\"
        Cells(currentRow, 9) = localFolder
    End If

    Call Fold

    Dim specialFile As String
    specialFile = Cells(currentRow, 11)

    If Trim(specialFile) = "" Or Not EndsWith(specialFile, ".sql") Or StartsWith(specialFile, "tmp_") Then
        specialFile = "tmp_" & Format(Now, "yyyyMMdd-HHmmss") & ".sql"
        Cells(currentRow, 11) = specialFile
    End If

    Dim filePath As String
    filePath = localFolder & specialFile

    WriteTxt2Tmp resultStr, filePath
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

End Sub

