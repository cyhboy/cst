
Public Sub DelVF()
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
                RobotRunByParam "DelVF"
            End If
        Next curCell
        Exit Sub
    End If

    Dim localFolder As String
    Dim fileName As String
    Dim orgFileName As String
    Dim videoFileName, audioFileName As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    localFolder = Cells(currentRow, 9)
    fileName = Cells(currentRow, 13)
    orgFileName = Left(fileName, InStrRev(fileName, ".") - 1)
    Dim fileList As Variant
    fileList = GetFileList(localFolder & orgFileName & ".f*")

    If UBound(fileList) < 2 Then
        Exit Sub
    End If

    Dim fileSize As Double
    fileSize = 1.79769313486231E+308

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myFileObj As Object
    Dim myFile As Variant

    For Each myFile In fileList
        Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
        'MsgBox myFileObj.Name & FileLen(localFolder & CStr(myFile))
        If FileLen(localFolder & CStr(myFile)) < fileSize Then
            videoFileName = audioFileName
            audioFileName = myFileObj.Name

            fileSize = FileLen(localFolder & CStr(myFile))
        End If
    Next myFile

    Set fso = Nothing

    MyQuestionBox "delete video file in row? " & videoFileName, "Yes", "No", 5
    If confirmation = "No" Then
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    'path = "cmd.exe /C C:\AppFiles\cmdutils\Recycle -f "
    path = "C:\AppFiles\cmdutils\Recycle.exe -f "
    'path = "Recycle.exe "

    parameter = """" & Cells(currentRow, 9) & videoFileName & """"

    ShellRun path & parameter, False

    Dim exeName As String: exeName = ExtractEXE(path)
    While True = IsExeRunning(exeName)
        Sleep 3000
    Wend

End Sub

