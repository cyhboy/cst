
Public Sub ExpTblWithoutPrompt()
    If testing Then
        Exit Sub
    End If
    'Dim lastrow As Long
    'lastrow = Range("B2").End(xlDown).Row
    Dim destFile As String
    Dim Suffix As String

    Suffix = Format(Now, "yyyyMMddhhmm")

    destFile = GetBakDrive() & "\" & ActiveSheet.Name & "_" & Suffix & ".txt"

    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select

    QuoteCommaExpByFileName destFile, 2, """"

    'MsgBox "DONE!"
End Sub

