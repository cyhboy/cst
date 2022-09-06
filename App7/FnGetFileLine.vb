
Public Function FnGetFileLine(filePath As String, theLineNo As Long) As String
    If testing Then
        Exit Function
    End If
    Const ForReading = 1, ForWriting = 2
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, ForReading)

    Dim lineNo As Long
    lineNo = 0

    Dim theLine As String

    Do Until ts.AtEndOfStream
        lineNo = lineNo + 1
        If lineNo = theLineNo Then
            theLine = ts.readline
            GoTo FoundHandler
        End If
    Loop
FoundHandler:
    Set ts = Nothing
    Set fso = Nothing
    FnGetFileLine = theLine
End Function

