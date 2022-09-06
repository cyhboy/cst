
Public Sub AccomplishEnvVar()
    If testing Then
        Exit Sub
    End If
    Dim parameter As String, sysParam As String
    Dim currentRow As Integer

    currentRow = ActiveCell.Row

    parameter = Cells(currentRow, 9)
    sysParam = Cells(currentRow, 3)

    Dim strQuery As String

    Dim strComputer As String
    strComputer = "."

    Dim objWMI As Object
    Set objWMI = GetObject("winmgmts://" & strComputer & "/root/cimv2")

    Dim colitems As Variant
    Dim objItem As Object
    If parameter = "" Then
        strQuery = "SELECT * FROM Win32_Environment"
        Set colitems = objWMI.ExecQuery(strQuery, "WQL", 48)
        Dim i As Integer
        i = 2
        For Each objItem In colitems
            ActiveSheet.Cells(i, 5).Value = objItem.Caption
            ActiveSheet.Cells(i, 6).Value = objItem.Description
            ActiveSheet.Cells(i, 3).Value = objItem.SystemVariable
            ActiveSheet.Cells(i, 4).Value = objItem.Status
            ActiveSheet.Cells(i, 1).Value = objItem.UserName

            ActiveSheet.Cells(i, 9).Value = objItem.Name
            ActiveSheet.Cells(i, 10).Value = Replace(objItem.VariableValue, ",", "|")

            i = i + 1
        Next objItem
    Else
        strQuery = "SELECT * FROM Win32_Environment WHERE SystemVariable=" & sysParam & " And Name='" & parameter & "'"
        'MsgBox strQuery
        Set colitems = objWMI.ExecQuery(strQuery)
        If colitems.count > 0 Then
            Cells(currentRow, 5).Value = colitems.ItemIndex(0).Caption
            Cells(currentRow, 6).Value = colitems.ItemIndex(0).Description
            Cells(currentRow, 3).Value = colitems.ItemIndex(0).SystemVariable
            Cells(currentRow, 4).Value = colitems.ItemIndex(0).Status
            Cells(currentRow, 1).Value = colitems.ItemIndex(0).UserName

            Cells(currentRow, 9).Value = colitems.ItemIndex(0).Name
            Cells(currentRow, 10).Value = Replace(colitems.ItemIndex(0).VariableValue, ",", "|")
        Else
            Cells(currentRow, 5).Value = ""
            Cells(currentRow, 6).Value = ""
            Cells(currentRow, 3).Value = ""
            Cells(currentRow, 4).Value = ""
            Cells(currentRow, 1).Value = "UNAVAILABLE"

            'Cells(currentRow, 9).Value = ""
            Cells(currentRow, 10).Value = "UNAVAILABLE"
        End If
    End If
End Sub

