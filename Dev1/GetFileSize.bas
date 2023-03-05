
Public Function GetFileSize(filename As String)
    If testing Then
        Exit Function
    End If
    On Error Resume Next
    Dim oFolder, ofPName As Variant
    With CreateObject("Shell.Application")
        Set oFolder = .Namespace(Left(filename, InStrRev(filename, "\") - 1))
        Set ofPName = oFolder.ParseName(Right(filename, Len(filename) - InStrRev(filename, "\")))
        GetFileSize = oFolder.GetDetailsOf(ofPName, 1)
    End With
End Function

