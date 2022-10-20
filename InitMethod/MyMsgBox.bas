
Public Sub MyMsgBox(detail As String, duration As Long)
    If testing Then
        Exit Sub
    End If

    nexttime = Now() + TimeSerial(0, 0, duration)
    Application.OnTime nexttime, "MyMsgBoxHide"

    'UserForm1.TextBox1.text = detail
    'UserForm1.TextBox1.SetFocus
    'UserForm1.Show
    Set uf1 = New UserForm1
    uf1.TextBox1.text = detail
    uf1.TextBox1.SetFocus
    uf1.Show
End Sub

