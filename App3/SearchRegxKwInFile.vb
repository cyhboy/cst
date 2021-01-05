
Public Function SearchRegxKwInFile(filePath As String, regxKw As String, Optional multiLine As Boolean = False, Optional ignoreC As Boolean = False)
    If testing Then Exit Function
    Dim fso, FileIn, strTmp
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set FileIn = fso.OpenTextFile(filePath, 1) 'for reading only
    
    Dim reg As New RegExp
    With reg
        .Global = True
        .IgnoreCase = ignoreC
        .multiLine = multiLine
        .Pattern = regxKw
    End With
    
    Dim mc As MatchCollection
    Dim dynamicStr1 As String
    
    
    If multiLine Then
        Dim strAll As String
        strAll = FileIn.readall
        If dynamicStr1 = "" And multiLine Then
            Set mc = reg.Execute(strAll)
            If mc.Count > 0 Then
                'MsgBox "hi"
                dynamicStr1 = mc.Item(0).SubMatches.Item(0)
            End If
        End If
    
    Else
        
        Do Until FileIn.AtEndOfStream
            strTmp = FileIn.readline
            If Len(strTmp) > 0 Then
                Set mc = reg.Execute(strTmp)
                If mc.Count > 0 Then
                    dynamicStr1 = mc.Item(0).SubMatches.Item(0)
                    Exit Do
                End If
            End If
        Loop
    
    End If
    
    FileIn.Close
    Set fso = Nothing
    
    SearchRegxKwInFile = dynamicStr1
End Function

