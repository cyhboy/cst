
Public Sub ExpRgn()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim destFile As String
    Dim Suffix As String
    Suffix = Format(Now, "yyyyMMddhhmm")

    destFile = GetBakDrive() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"

    Range("A1").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    'Range(Selection, Selection.End(xlDown)).Select

    ActiveCell.CurrentRegion.Select
    QuoteCommaExpByFileName destFile, 1, """"
    MyMsgBox "DONE!", 2

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

End Sub

