
Public Sub ExpTbl()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    If Range("A2") = "" Then
        Exit Sub
    End If
    Dim destFile As String
    Dim Suffix As String
    Suffix = Format(Now, "yyyyMMddhhmm")

    destFile = GetBakDrive() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    'Range("A1").CurrentRegion.Select

    QuoteCommaExpByFileName destFile, 2, """"

    MyMsgBox "Done", 10

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

