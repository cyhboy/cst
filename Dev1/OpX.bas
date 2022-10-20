
Public Sub OpX()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    
    Dim album As String
    Dim artist As String
    Dim path As String
    path = Cells(2, 9)
    
    If InStr(path, "\") <= 0 Then
        Exit Sub
    End If
    
    album = Split(path, "\")(UBound(Split(path, "\")) - 1)
    artist = Split(path, "\")(UBound(Split(path, "\")) - 2)

    Dim midStr As String
    midStr = ""
    Dim rplStr As String

    Dim findOut As Range
    Set findOut = Cells.Find(What:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False)

    If Not findOut Is Nothing Then
        MyQuestionBox "keep original video/audio file or not?", "Yes", "No", 5
        If confirmation = "Yes" Then
            midStr = midStr & "-k "
        End If

        MyQuestionBox "how about generate audio file after download?", "Yes", "No", 5
        If confirmation = "Yes" Then
            midStr = midStr & "-x --audio-format flac "
            MyQuestionBox "how about apply compression to all audio file?", "Yes", "No", 5
            If confirmation = "Yes" Then
                midStr = Replace(midStr, "--audio-format flac ", "")
            End If
            MyQuestionBox "keep metadata to all audio file?", "Yes", "No", 5
            If confirmation = "Yes" Then
                'midStr = midStr & "--add-metadata --postprocessor-args ""-metadata album=Level2LearnEnglishthroughStory"" "
                midStr = midStr & "--postprocessor-args ""-metadata album=" & album & " -metadata artist=" & artist & """ "
            End If
        End If
        rplStr = "youtube-dl " & midStr & "--cookies"
        
        Cells.Replace What:="youtube-dl --cookies", Replacement:=rplStr, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Else
        Cells.Replace What:="youtube-dl * --cookies", Replacement:="youtube-dl --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    End If
End Sub


