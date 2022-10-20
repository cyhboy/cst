
Public Sub L2Y()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="C:\lingo\", Replacement:="C:\Users\cyy\Desktop\youtube\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

