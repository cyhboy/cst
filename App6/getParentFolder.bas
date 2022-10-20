
Public Function getParentFolder(strFolder As String)
    If testing Then
        Exit Function
    End If

    If EndsWith(strFolder, "\") Then
        strFolder = Left(strFolder, Len(strFolder) - 1)
    End If
    getParentFolder = Left(strFolder, InStrRev(strFolder, "\") - 1)
End Function

