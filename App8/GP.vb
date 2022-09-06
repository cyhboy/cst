
Public Sub GP()
    If testing Then
        Exit Sub
    End If
    Dim videoPath As String
    Dim videoFileName As String
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    videoPath = Cells(currentRow, 9)
    videoFileName = Cells(currentRow, 11)
    Dim FullPath As String
    FullPath = videoPath & videoFileName

    Dim exeName As String: exeName = ExtractEXE("dllhost.exe")
    Dim cntEXE As Integer
    cntEXE = CntExeRunning(exeName)

    With CreateObject("Shell.Application").Namespace(0).ParseName(FullPath)
        .Invokeverb "Properties"
    End With

    While CntExeRunning(exeName) = cntEXE + 1
        Sleep 3000
    Wend
End Sub


