
Public Sub VD()
    If testing Then
        Exit Sub
    End If
    Dim videoPath As String
    Dim videoFileName As String
    Dim videoFullFilename As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    videoPath = Cells(currentRow, 9)
    videoFileName = Cells(currentRow, 11)

    If Trim(videoPath) = "" Or Trim(videoFileName) = "" Then
        Exit Sub
    End If
    videoFullFilename = videoPath & videoFileName

    Cells(currentRow, 7) = GetVideoDuration(videoFullFilename)
End Sub

