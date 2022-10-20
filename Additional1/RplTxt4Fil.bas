
Public Sub RplTxt4Fil(sFileName As String, orgTxt As String, newTxt As String)
    If testing Then
        Exit Sub
    End If
    'Call by RplTxt4Fld
    Dim sBuf As String
    Dim sTemp As String
    Dim ff As Integer

    ff = FreeFile
    Open sFileName For Input As #ff

    Do Until EOF(ff)
        Line Input #ff, sBuf
        sTemp = sTemp & sBuf & vbCrLf
    Loop

    sTemp = Left(sTemp, Len(sTemp) - Len(vbCrLf))

    Close #ff

    sTemp = Replace(sTemp, orgTxt, newTxt)

    ff = FreeFile
    Open sFileName For Output As #ff
    Print #ff, sTemp
    Close #ff

End Sub

