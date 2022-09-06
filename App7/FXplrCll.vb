
Public Sub FXplrCll()
    If testing Then
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    Dim count As Integer
    count = 0

    Dim cell As Object
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            count = count + 1
        End If
    Next cell

    If count Mod 2 <> 0 Then
        'Selection.Cells.Rows.count & Selection.Cells.Columns.count
        MsgBox "Please let the cell size be pair!"
        Exit Sub
    End If

    Dim path As String
    Dim parameter As String
    path = "explorer "

    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            If InStr(cell.Value, "\") > 0 Then
                If EndsWith(cell.Value, "\") Then
                    parameter = cell.Value & parameter
                Else
                    parameter = cell.Value & "\" & parameter
                End If
            Else
                parameter = parameter & cell.Value
            End If
        End If
    Next cell
    parameter = """" & parameter & """"
    ShellRun path & parameter, False
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub

