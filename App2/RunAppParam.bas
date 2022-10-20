
Public Sub RunAppParam(isHold As Boolean, isTest As Boolean, isKeep As Boolean)
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim parameter As String

    parameter = Cells(currentRow, 10)

    Dim arr As Variant

    arr = Split(parameter, Chr(10))

    Dim path As String
    Dim i As Integer
    For i = 0 To UBound(arr)
        If Not (StartsWith(Trim(arr(i)), "::") Or StartsWith(UCase(Trim(arr(i))), "REM")) Then
            path = path & arr(i) & "&"
        End If
    Next i

    While Right(path, 1) = "&"
        path = Left(path, Len(path) - 1)
    Wend

    If Not Cells(currentRow, 9).HasFormula Then
        If Dir(Cells(currentRow, 9), vbDirectory) <> vbNullString Then
            Dim cdPath As String
            cdPath = Cells(currentRow, 9)
            path = "cd " & cdPath & "&" & path
        End If
    End If

    If isTest Then
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")
        path = "PsExec -u GUANGZHOUTEST\" & Environ$("username") & " -p " & propsMap("AD_PASSWORD_UAT") & " " & path
        'MsgBox path
        ShellRunHide path
    Else
        ShellRun path, isKeep
    End If

    If isHold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 5000
        Wend
    End If
End Sub

