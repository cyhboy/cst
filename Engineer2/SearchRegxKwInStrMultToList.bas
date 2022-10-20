
Public Function SearchRegxKwInStrMultToList(str As String, regxKw As String, matchI As Integer, multiFlag As Boolean)
    If testing Then
        Exit Function
    End If

    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = False
        .multiLine = multiFlag
        .Pattern = regxKw
    End With

    Dim mc As MatchCollection
    'Dim dynamicStr1 As String

    Set mc = reg.Execute(str)

    If mc.count > 0 Then
        ReDim strArr(mc.count - 1) As String
    End If

    Dim i As Integer
    If mc.count > 0 Then
        For i = 0 To mc.count - 1
            strArr(i) = Replace(mc.Item(i).SubMatches.Item(matchI), ",", ";")
        Next i
    End If

    SearchRegxKwInStrMultToList = strArr
End Function

