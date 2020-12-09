
Private Sub CommandButton1_Click()
    Application.OnTime nexttime, "MyQuestionBoxHide", , False
    'UserForm2.Hide
    'confirmation = UserForm2.CommandButton1.Caption
    uf2.Hide
    confirmation = uf2.CommandButton1.Caption
    Set uf2 = Nothing
End Sub

