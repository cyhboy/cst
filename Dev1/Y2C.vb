
Public Sub Y2C()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="C:\Users\cyy\Desktop\youtube\", Replacement:="C:\choice\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

