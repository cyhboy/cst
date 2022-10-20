
Public Function SearchRegxKwInStr(str As String, regxKw As String, Optional multiLine As Boolean = False, Optional ignoreC As Boolean = False)
    'SearchRegxKwInStr
    If testing Then
        Exit Function
    End If
    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = ignoreC
        .multiLine = multiLine
        .Pattern = regxKw
    End With

    Dim mc As MatchCollection
    Dim dynamicStr1 As String
    dynamicStr1 = ""
    Set mc = reg.Execute(str)
    If mc.count > 0 Then
        dynamicStr1 = mc.Item(0).SubMatches.Item(0)
    End If

    SearchRegxKwInStr = dynamicStr1
End Function

