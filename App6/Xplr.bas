
Public Sub Xplr()
    If testing Then
        Exit Sub
    End If

    Dim path As String
    Dim folderVal As String
    Dim parameter As String
    path = "explorer "

    Dim cell As Object
    Dim currentRow As Integer
    'Dim fso As Object
    'Set fso = CreateObject("Scripting.FileSystemObject")
    For Each cell In Selection.Cells
        If cell.EntireColumn.Hidden = False And cell.EntireRow.Hidden = False Then
            currentRow = cell.Row
            folderVal = Cells(currentRow, 9)

            If Not Dir(folderVal, vbDirectory) = "" Then
                parameter = """" & folderVal & """"
            Else
                MyMsgBox "will open parent folder as current path is not existing!", 3
                parameter = """" & getParentFolder(folderVal) & """"
            End If
            ShellRun path & parameter, False
        End If
    Next cell
    'Set fso = Nothing
End Sub

