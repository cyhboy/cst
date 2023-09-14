
Public Sub YTDLP()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="youtube-dl", Replacement:="yt-dlp", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
    Cells.Replace What:="yt-dlp --cookies", Replacement:="yt-dlp --proxy ""socks5://127.0.0.1:7890"" --cookies", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

