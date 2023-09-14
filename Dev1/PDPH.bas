
Public Sub PDPH()
    ' pandas post handler
    If testing Then
        Exit Sub
    End If
    Call UnHF
    Cells.Replace What:="_x000D_", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, ReplaceFormat:=False
End Sub

