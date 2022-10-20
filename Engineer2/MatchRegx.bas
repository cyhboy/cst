
Public Function MatchRegx(text As String, patt As String, Optional ignoreC As Boolean = False) As Boolean
    If testing Then
        Exit Function
    End If
    'Set up regular expression object
    Dim RE As New RegExp
    RE.Pattern = patt
    RE.Global = True
    RE.IgnoreCase = ignoreC
    RE.multiLine = True
    'Retrieve all matches
    Dim Matches As MatchCollection
    Set Matches = RE.Execute(text)
    'Return the corrected count of matches
    If Matches.count > 0 Then
        MatchRegx = True
    Else
        MatchRegx = False
    End If
End Function

