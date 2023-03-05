
Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String) As String
    If testing Then
        Exit Function
    End If

    Dim Worked As Long
    Dim RetStr As String * 128
    Dim StrSize As Long

    Dim iNoOfCharInIni As Integer: iNoOfCharInIni = 0
    Dim sIniString As String: sIniString = ""
    Dim sProfileString As String
    
    If Sect = "" Or Keyname = "" Then
        MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
    Else
        sProfileString = ""
        RetStr = Space(128)
        StrSize = Len(RetStr)
        Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
        If Worked Then
            iNoOfCharInIni = Worked
            sIniString = Left$(RetStr, Worked)
        End If
    End If
    ReadIniFileString = sIniString
End Function

