
Public Function FnLstFil(fileSpec As String) As String
    If testing Then
        Exit Function
    End If
    Dim filePath As String
    Dim fileList As Variant
    fileList = GetFileList(fileSpec)

    Dim localFolder As String
    localFolder = Left(fileSpec, InStrRev(fileSpec, "\"))
    'MsgBox localFolder
    Dim date1 As Date

    date1 = DateAdd("yyyy", -5, Now)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myFileObj As Object
    Dim myFile As Variant
    For Each myFile In fileList
        Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
        If myFileObj.DateLastModified > date1 Then
            date1 = myFileObj.DateLastModified
            'If rtnType = 1 Then filename1 = myFile.path
            'If rtnType = 2 Then filename1 = myFile.Name
            filePath = myFileObj.path
        End If
    Next myFile
    Set fso = Nothing
    FnLstFil = filePath
End Function


