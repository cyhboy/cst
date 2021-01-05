
Public Function ReadLineByFile(fileName As String)
    If testing Then Exit Function
    On Error GoTo ErrorHandler
    Const ForReading = 1, ForWriting = 2
    
    Dim fso, fro As Object
    Dim readResult As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fro = fso.OpenTextFile(fileName, ForReading)

    readResult = fro.readall
    
    Set fro = Nothing
    Set fso = Nothing
    
    ReadLineByFile = Trim(readResult)
    
ErrorHandler:
    If Err.Number = 62 Then
        ReadLineByFile = ""
        Exit Function
    End If
    
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Function

