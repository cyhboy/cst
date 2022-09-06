
Public Function GetVideoDuration(fileName As String)
    If testing Then
        Exit Function
    End If
    On Error GoTo ErrorHandler
    Dim oFolder, ofPName As Variant
    With CreateObject("Shell.Application")
        Set oFolder = .Namespace(Left(fileName, InStrRev(fileName, "\") - 1))
        Set ofPName = oFolder.ParseName(Right(fileName, Len(fileName) - InStrRev(fileName, "\")))
        GetVideoDuration = CDbl(TimeValue(oFolder.GetDetailsOf(ofPName, 27))) * 24# * 60#
    End With
ErrorHandler:
    If Err.Number <> 0 Then
        GetVideoDuration = ""
    End If
End Function

