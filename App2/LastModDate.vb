
Public Function LastModDate(filePath As String) As Date
    If testing Then
        Exit Function
    End If
    On Error GoTo ErrorHandler
    Dim resultDate As Date
    'resultDate = DateAdd("yyyy", -7, Now)
    resultDate = Date

    Dim fso As Object
    Dim fileObject As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If EndsWith(filePath, "\") Then
        Set fileObject = fso.getfolder(filePath)
    Else
        Set fileObject = fso.GetFile(filePath)
    End If

    resultDate = fileObject.DateLastModified()
    Set fso = Nothing

ErrorHandler:
    LastModDate = resultDate
End Function

