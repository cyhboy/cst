
Private Sub CommandButton2_Click()
    Application.OnTime nexttime, "MyQuestionBoxHide", , False
    'UserForm2.Hide
    'confirmation = UserForm2.CommandButton2.Caption
    uf2.Hide
    confirmation = uf2.CommandButton2.Caption
    Set uf2 = Nothing
End Sub



