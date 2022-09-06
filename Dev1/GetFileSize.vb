
Public Function GetFileSize(fileName As String)
    If testing Then
        Exit Function
    End If
    On Error Resume Next
    Dim oFolder, ofPName As Variant
    With CreateObject("Shell.Application")
        Set oFolder = .Namespace(Left(fileName, InStrRev(fileName, "\") - 1))
        Set ofPName = oFolder.ParseName(Right(fileName, Len(fileName) - InStrRev(fileName, "\")))
        GetFileSize = oFolder.GetDetailsOf(ofPName, 1)
    End With
End Function

