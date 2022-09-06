
Public Sub OpX()
    If testing Then
        Exit Sub
    End If
    Call UnHF

    Dim rplStr As String

    Dim findOut As Range
    Set findOut = Cells.Find(What:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)

    If Not findOut Is Nothing Then
        MyQuestionBox "keep original video/audio file or not?", "Yes", "No", 5
        If confirmation = "Yes" Then
            rplStr = "youtube-dl -k --cookies"
        End If

        MyQuestionBox "how about generate audio file after download?", "Yes", "No", 5
        If confirmation = "Yes" Then
            rplStr = Replace(rplStr, " --cookies", " -x --audio-format flac --cookies")
            MyQuestionBox "how about apply compression to all audio file?", "Yes", "No", 5
            If confirmation = "Yes" Then
                rplStr = Replace(rplStr, "--audio-format flac ", "")
            End If
        End If

        Cells.Replace What:="youtube-dl --cookies", Replacement:=rplStr, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Else
        Cells.Replace What:="youtube-dl * --cookies", Replacement:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    End If
End Sub


