
Public Sub MyQuestionBoxHide()
    If testing Then Exit Sub
    'confirmation = UserForm2.CommandButton1.Caption
    'UserForm2.Hide
    confirmation = uf2.CommandButton1.Caption
    uf2.Hide
    Set uf2 = Nothing
    'MsgBox "This is a scheduler"
End Sub

