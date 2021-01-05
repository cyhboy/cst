
Public Sub RunAppParam(isHold As Boolean, isTest As Boolean)
    If testing Then Exit Sub
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim parameter As String
    
    parameter = Cells(currentRow, 10)
    
    Dim arr
    
    arr = Split(parameter, Chr(10))
    
    Dim path As String
    Dim i
    For i = 0 To UBound(arr)
        path = path & arr(i) & "&"
    Next
    
    While Right(path, 1) = "&"
        path = Left(path, Len(path) - 1)
    Wend
    
    If isTest Then
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        path = "PsExec -u GUANGZHOUTEST\" & Environ$("username") & " -p " & propsMap("AD_PASSWORD_UAT") & " " & path
        MsgBox path
        ShellRunHide path
    Else
        ShellRun path
    End If
    
    If isHold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 5000
        Wend
    End If
End Sub

