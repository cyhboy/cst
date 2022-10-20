
Public Sub CTY()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="C:\choice\", Replacement:="C:\Users\cyy\Desktop\youtube\", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

