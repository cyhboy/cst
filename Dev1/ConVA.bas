
Public Sub ConVA()
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
                RobotRunByParam "ConVA"
            End If
        Next curCell
        Exit Sub
    End If

    Dim localFolder As String
    Dim filename As String
    Dim orgFileName As String
    Dim videoFileName, audioFileName As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    localFolder = Cells(currentRow, 9)
    filename = Cells(currentRow, 13)
    orgFileName = Left(filename, InStrRev(filename, ".") - 1)
    Dim fileList As Variant
    fileList = GetFileList(localFolder & orgFileName & ".f*")

    If UBound(fileList) < 2 Then
        Exit Sub
    End If

    Dim fileSize As Double
    fileSize = 0

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myFileObj As Object
    Dim myFile As Variant

    For Each myFile In fileList
        Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
        If FileLen(localFolder & CStr(myFile)) > fileSize Then
            audioFileName = videoFileName
            videoFileName = myFileObj.Name
            fileSize = FileLen(localFolder & CStr(myFile))
        Else
            videoFileName = audioFileName
            audioFileName = myFileObj.Name
        End If
    Next myFile
    Set fso = Nothing

    Dim cmdStr As String

    ' MsgBox GetVideoDuration(localFolder & videoFileName)
    ' MsgBox IsNumeric(GetVideoDuration(localFolder & videoFileName))
    If IsNumeric(GetVideoDuration(localFolder & videoFileName)) And 1 <> 1 Then
        ' Never go here as I found copy merge also can fail on a normal case
        cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -c copy """ & localFolder & filename & """"
    Else
        ' cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -c copy -map 0:v -map 1:a -shortest -af apad """ & localFolder & fileName & """"
        cmdStr = "ffmpeg -y -i """ & localFolder & videoFileName & """ -i """ & localFolder & audioFileName & """ -af apad -shortest """ & localFolder & filename & """"
    End If
    ' MsgBox cmdStr
    ShellRun cmdStr, True

End Sub

