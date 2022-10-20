
Public Sub ImpCsvLegacy()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False
    Call CleanRgn

    Dim filePath As String
    filePath = FnLstFil("C:\BAK\*.csv")

    Dim ff As Integer
    ff = FreeFile()

    Open filePath For Input As #ff
    Dim rowNo As Integer
    Dim lineStr As String
    Dim lineItems As Variant
    rowNo = 0
    Dim colNo As Integer
    Dim cllVal As String
    Do Until EOF(ff)
        Line Input #ff, lineStr
        If InStr(lineStr, "'") > 0 Then
            'lineStr = Left(lineStr, Len(lineStr) - 1)
            'lineStr = Right(lineStr, Len(lineStr) - 1)
            'lineItems = Split(lineStr, "','")
            lineItems = Split(lineStr, ",")
        Else
            lineItems = Split(lineStr, ",")
        End If

        For colNo = 0 To UBound(lineItems)
            cllVal = lineItems(colNo)
            If StartsWith(cllVal, "'") And EndsWith(cllVal, "'") Then
                cllVal = Left(cllVal, Len(cllVal) - 1)
                cllVal = Right(cllVal, Len(cllVal) - 1)
            End If

            If IsNumeric(cllVal) Then
                cllVal = "'" & cllVal
            End If

            Cells(rowNo + 1, colNo + 1).Value = cllVal
        Next colNo
        rowNo = rowNo + 1
    Loop
    Close #ff
    Application.ScreenUpdating = True
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

