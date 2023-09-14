
Public Function ShellRunResult(cmd As String, Optional cmdLogFile As String = "C:\BAK\cmd.log", Optional Hold As Boolean = True)
    If testing Then
        Exit Function
    End If

    'Dim cmdLogFile As String
    'cmdLogFile = "C:\BAK\cmd.log"
    Dim path As String
    path = "cmd.exe /C " & cmd & " > " & cmdLogFile

    path = Replace(path, "2 >", "2>")

    Dim cntEXE As Integer
    cntEXE = CntExeRunning(ExtractEXE(path))

    If InStr(path, " npm ") > 0 Then
        cntEXE = cntEXE + CntExeRunning("node.exe")
    End If

    'MsgBox path
    'Exit Function
    
    Shell path, vbNormalFocus
    'Shell path, vbHide

    If Hold Then
        While Hold
            Dim cntEXE2 As Integer
            cntEXE2 = CntExeRunning(ExtractEXE(path))
            If InStr(path, " npm ") > 0 Then
                cntEXE2 = cntEXE2 + CntExeRunning("node.exe")
            End If
            'MsgBox cntEXE2
            If cntEXE2 - cntEXE > 0 Then
                Sleep 3000
            Else
                Hold = False
            End If
        Wend
        'If cntEXE = 0 Then
        '    Sleep 3000
        'End If

        'If CntExeRunning(ExtractEXE(path)) = 0 Then
        '    Sleep 3000
        'End If
    End If

    ShellRunResult = ReadLineByFile(cmdLogFile)

End Function

