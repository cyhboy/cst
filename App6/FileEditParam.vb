
Public Sub FileEditParam(hold As Boolean, isFilter As Boolean, path As String)
    If testing Then
        Exit Sub
    End If

    Dim parameter As String

    'path = """" & GetAppDrive() & "\EditPlus\editplus.exe" & """ -e"
    'path = """" & "C:\Program Files\IDM Computer Solutions\UltraEdit\Uedit32.exe" & """ "
    'path = "C:\AppFiles\Microsoft VS Code\Code.exe"

    'path = "C:\AppFiles\SublimeText\sublime_text.exe"
    'path = "C:\AppFiles\Notepad++\notepad++.exe"
    'path = "C:\Program Files\Microsoft VS Code\bin\code.cmd"
    'path = "code.cmd"

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    'parameter = " " & """" & Replace(cell.Value, Chr(10), """ """) & """"
    Dim fileStr As String
    If Not isFilter Then
        fileStr = Cells(currentRow, 11)
    Else
        fileStr = "*"
    End If

    Dim fileArr As Variant
    fileArr = Split(fileStr, Chr(10))
    Dim i As Integer
    For i = 0 To UBound(fileArr)
        parameter = " " & """" & Cells(currentRow, 9) & fileArr(i) & """"
        ' MsgBox path & parameter
        ' ShellRunHide path & parameter
        ShellRun path & parameter, False
    Next i
    If hold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 5000
        Wend
    End If
End Sub

