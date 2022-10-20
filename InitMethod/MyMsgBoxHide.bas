
Public Sub MyMsgBoxHide()
    If testing Then
        Exit Sub
    End If

    'UserForm1.Hide
    uf1.Hide
    Set uf1 = Nothing
    'MsgBox "This is a scheduler"
End Sub

