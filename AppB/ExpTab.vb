
Public Sub ExpTab()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    Dim destFile As String
    Dim Suffix As String
    Suffix = Format(Now, "yyyyMMddhhmmss")
    destFile = GetBakDrive() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"

    If Range("A2") = "" Then
        WriteTxt2Tmp "", destFile
        Exit Sub
    End If

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    'Range("A1").CurrentRegion.Select

    QuoteTabExpByFileName destFile, 1

    MyMsgBox "Done", 3

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

