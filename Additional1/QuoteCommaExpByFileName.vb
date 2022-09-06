
Public Sub QuoteCommaExpByFileName(destFile As String, firstLineNo As Long, quote As String)
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim ff As Integer
    Dim ColumnCount As Integer
    Dim RowCount As Long

    ff = FreeFile()

    ' Close Error Handling

    ' Try to open target file for output
    Open destFile For Output As #ff
    '
    If Err <> 0 Then
        MsgBox "Cannot open filename " & destFile
    End If

    ' Open Error Handling

    Dim idx As Integer
    idx = ActiveSheet.index - 1
    'MsgBox Idx

    '
    For RowCount = firstLineNo To Selection.Rows.count

        If Selection.Rows(RowCount).Hidden = False Then
            '
            For ColumnCount = 1 To Selection.Columns.count

                If Selection.Columns(ColumnCount).Hidden = False Then
                    '
                    Dim valueType As String

                    If idx <> 0 Then
                        valueType = Sheets(idx).Cells(ColumnCount + 1, 6).Value
                        If (valueType = "INTEGER" Or valueType = "SMALLINT" Or valueType = "BIGINT" Or valueType = "DECIMAL") Then
                            Print #ff, Selection.Cells(RowCount, ColumnCount).Value;
                        Else
                            'MsgBox Replace(Selection.Cells(RowCount, ColumnCount).Value, """", "")
                            Dim valueAfterReplace As String
                            'MsgBox Selection.Cells(RowCount, ColumnCount).Value
                            valueAfterReplace = Replace(Replace(Replace(Selection.Cells(RowCount, ColumnCount).Value, quote, "\" & quote), Chr(10), "\n"), Chr(13), "")

                            If ((valueType = "TIMESTAMP" Or valueType = "DATE") And valueAfterReplace = "") Then
                                'Let date/time/timestamp be null
                                Print #ff, valueAfterReplace;
                                'Print #ff, """" & valueAfterReplace & """";
                            Else
                                If (IsNumeric(valueAfterReplace) And InStr(valueAfterReplace, ".") > 0) Or valueAfterReplace = "getdate()" Or valueAfterReplace = "NULL" Then
                                    Print #ff, valueAfterReplace;
                                Else
                                    Print #ff, quote & valueAfterReplace & quote;
                                End If
                            End If
                        End If
                    Else
                        If (IsNumeric(Selection.Cells(RowCount, ColumnCount).Value) And InStr(Selection.Cells(RowCount, ColumnCount).Value, ".") > 0) Or Selection.Cells(RowCount, ColumnCount).Value = "getdate()" Or Selection.Cells(RowCount, ColumnCount).Value = "NULL" Then
                            Print #ff, Selection.Cells(RowCount, ColumnCount).Value;
                        Else
                            Print #ff, quote & Selection.Cells(RowCount, ColumnCount).Value & quote;
                        End If
                    End If

                    If ColumnCount = Selection.Columns.count Then
                        Print #ff,
                    Else
                        Print #ff, ",";
                    End If
                End If

            Next ColumnCount
        End If

    Next RowCount

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
    ' Close Target File
    Close #ff
End Sub

