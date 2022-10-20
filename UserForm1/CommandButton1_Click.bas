
Private Sub CommandButton1_Click()
    Application.OnTime nexttime, "MyMsgBoxHide", , False
    'UserForm1.Hide
    uf1.Hide
    Set uf1 = Nothing
End Sub



