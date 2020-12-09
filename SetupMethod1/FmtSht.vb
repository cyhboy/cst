
Public Sub FmtSht()
    If testing Then Exit Sub
    Call CreTitl
    Range("A2").Select
    Call FitScr
    Call Sample
    Call DrawTbl
    Call Frz
    Call RstCf
End Sub

