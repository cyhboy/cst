
Public Function CountRegx(text As String, patt As String) As Long
'    If testing Then
'        Exit Function
'    End If
    On Error GoTo ErrorHandler
    Dim RE As New RegExp
    RE.Pattern = patt
    RE.Global = True
    RE.IgnoreCase = False
    RE.multiLine = True
    'Retrieve all matches
    Dim Matches As MatchCollection
    Set Matches = RE.Execute(text)
    'Return the corrected count of matches
    CountRegx = Matches.count
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Function

