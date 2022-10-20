
Public Sub CmpFil()
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
    path = """" & GetAppDrive() & "\Beyond Compare 3\BCompare.exe" & """"
    '/fv=""Text Compare""
    Dim i As Integer
    i = 1
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            If i Mod 2 = 0 Then
                parameter = parameter & " " & """" & cell.Value & """"
                ShellRun path & parameter, False
            Else
                parameter = " " & """" & cell.Value & """"
            End If

            i = i + 1
        End If
    Next cell

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

