
Private Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String) As String
    If testing Then
        Exit Function
    End If

    Dim Worked As Long
    Dim iNoOfCharInIni As Integer: iNoOfCharInIni = 0
    Dim sIniString As String: sIniString = ""
    If Sect = "" Or Keyname = "" Then
        MsgBox "Section Or Key To Write Not Specified !!!", vbExclamation, "INI"
    Else
        Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
        If Worked Then
            iNoOfCharInIni = Worked
            sIniString = Wstr
        End If
        WriteIniFileString = sIniString
    End If
End Function

