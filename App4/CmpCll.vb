
Public Sub CmpCll()
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

    If count <> 2 Then
        'Selection.Cells.Rows.count & Selection.Cells.Columns.count
        MsgBox "Please let the selected cell size be 2!"
        Exit Sub
    End If

    Dim i As Integer
    i = 1
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            If i Mod 2 = 0 Then
                WriteTxt2Tmp cell.Value, GetBakDrive() & "\tmp2.txt"
            Else
                WriteTxt2Tmp cell.Value, GetBakDrive() & "\tmp1.txt"
            End If
            i = i + 1
        End If
    Next cell

    Dim path As String
    Dim parameter As String
    path = """" & GetAppDrive() & "\Beyond Compare 3\BCompare.exe" & """"
    parameter = " " & """" & GetBakDrive() & "\tmp1.txt" & """" & " " & """" & GetBakDrive() & "\tmp2.txt" & """"
    'MsgBox path & parameter
    ShellRun path & parameter, False

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

