
Public Function GetFileList(fileSpec As String) As Variant
    If testing Then
        Exit Function
    End If
    '   Returns an array of filenames that match FileSpec
    '   If no matching files are found, it returns False

    Dim FileArray() As Variant
    Dim FileCount As Integer
    Dim fileName As String

    On Error GoTo NoFilesFound

    FileCount = 0
    fileName = Dir(fileSpec)
    ' MsgBox fileSpec
    ' MsgBox fileName
    If fileName = "" Then GoTo NoFilesFound

        '   Loop until no more matching files are found
        Do While fileName <> ""
            FileCount = FileCount + 1
            ReDim Preserve FileArray(1 To FileCount)
            FileArray(FileCount) = fileName
            fileName = Dir()
            'MsgBox FileName
        Loop

        GetFileList = FileArray
        Exit Function

        '   Error handler
NoFilesFound:
        ReDim FileArray(0)
        GetFileList = FileArray
End Function

