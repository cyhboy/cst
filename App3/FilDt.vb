
Public Sub FilDt()
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
                RobotRunByParam "FilDt"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim localFolder As String
    Dim specialFile As String
    localFolder = Cells(currentRow, 9)
    specialFile = Cells(currentRow, 11)
    Dim filePath As String
    filePath = localFolder & specialFile
    Dim modifyDate As Date
    modifyDate = DateAdd("yyyy", -15, Now)


    If InStr(filePath, "*") = 0 And InStr(filePath, "?") = 0 Then
        modifyDate = LastModDate(filePath)

    Else
        Dim lstFilename As String
        Dim fileList As Variant
        fileList = GetFileList(filePath)

        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim myFileObj As Object
        Dim myFile As Variant
        For Each myFile In fileList
            'If CStr(myFile) <> currFilename Then
            Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
            If myFileObj.DateLastModified > modifyDate Then
                modifyDate = myFileObj.DateLastModified
                'If rtnType = 1 Then filename1 = myFile.path
                'If rtnType = 2 Then filename1 = myFile.Name
                lstFilename = myFileObj.Name
            End If
            'End If
        Next myFile
        Set fso = Nothing
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    Cells(currentRow, 12) = modifyDate
End Sub

