
Public Sub CreateFolder(path As String)
    If testing Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then
        Exit Sub
    End If
    
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then
        CreateFolder fso.GetParentFolderName(path)
    End If
    fso.CreateFolder path
    Set fso = Nothing
End Sub

