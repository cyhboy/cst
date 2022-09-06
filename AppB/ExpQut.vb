
Public Sub ExpQut()
    If testing Then
        Exit Sub
    End If
    '    On Error GoTo ErrorHandler
    If Range("A1") = "" Then
        Exit Sub
    End If
    Dim destFile As String
    Dim Suffix As String
    Suffix = Format(Now, "yyyyMMddhhmm")

    destFile = GetBakDrive() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"
    'MsgBox destFile
    Range("A1").Select

    If Trim(Range("B1")) <> "" Then
        Range(Selection, Selection.End(xlToRight)).Select
    End If

    If Trim(Range("A2")) <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
    End If
    'Range("A1").CurrentRegion.Select

    QuoteCommaExpByFileName destFile, 1, "'"

    MyMsgBox "Done", 10

    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub

