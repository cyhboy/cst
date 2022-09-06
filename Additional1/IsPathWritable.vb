
Public Function IsPathWritable(ByVal fPath As String) As String
    If testing Then
        Exit Function
    End If
    Dim fName As String
    'Dim localFName As String
    Dim ff As Integer
    Dim Counter As Integer
    Counter = 1

    If (Right(fPath, 1) <> "\") Then
        fPath = fPath & "\"
    End If

    IsPathWritable = "Invalid"

    On Error GoTo ErrHandler

    Do
        fName = fPath & "TempFile" & Counter & ".tmp"
        'localFName = "C:\Temp\" & "TempFile" & Counter & ".tmp"
        Counter = Counter + 1
    Loop Until Dir(fName) = ""

    On Error GoTo CantWrite

    ff = FreeFile()
    Open fName For Output Access Write As #ff
    Print #ff, "TESTWRITE"
    Close #ff


    'FileCopy FName, localFName
    'FileCopy localFName, FName

    On Error GoTo CantDelete
    Kill fName
    IsPathWritable = "Modifiable"
    Exit Function

CantDelete:
    IsPathWritable = "Writeable"
    Exit Function

CantWrite:
    IsPathWritable = "Readable"
    Exit Function

ErrHandler:
    Exit Function
End Function

