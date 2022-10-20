
Public Sub Y2L()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="C:\Users\cyy\Desktop\youtube\", Replacement:="C:\lingo\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

