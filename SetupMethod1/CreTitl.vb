
Public Sub CreTitl()
    If testing Then
        Exit Sub
    End If
    Dim titlAry As Variant
    titlAry = Array("Hostname", "FQDN", "User", "Password", "Folder", "IP", "Port", "Memo", "Local Folder", "Command", "Specify File", "Last Update", "Demand", "CO", "Sequence", "Executor", "Status", "#", "#", "#", "#", "#", "#")
    Dim i As Integer
    For i = 0 To UBound(titlAry)
        If Trim(Cells(1, i + 1)) = "" Then
            Cells(1, i + 1) = titlAry(i)
        End If
        Next
End Sub

