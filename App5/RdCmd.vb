
Public Sub RdCmd()
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
                RobotRunByParam "RdCmd"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row


    Dim localPath As String
    localPath = Cells(currentRow, 9)
    'localPath = Replace(localPath, "\", "/")
    Dim wildcard As String
    'wildcard = Cells(currentRow, 11)
    wildcard = "*"

    Dim fileText As String

    Dim fso As Object
    Dim objFolder As Object
    Dim myFile As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(localPath) Then
        Set objFolder = fso.getfolder(localPath)

        For Each myFile In objFolder.Files
            If (myFile.Name Like wildcard Or myFile.Name = wildcard) And (Not myFile.Name Like "tmp_*") Then
                fileText = fileText & ReadLineByFile(myFile.path)
            End If
        Next myFile
        Set objFolder = Nothing
    End If
    Set fso = Nothing


    Cells(currentRow, 10) = "'" & fileText
ErrorHandler:
    If Err.Number <> 0 Then
        'PrintResult "FAILED"
        MyMsgBox Err.Number & " " & Err.Description, 10
    Else

    End If
End Sub

