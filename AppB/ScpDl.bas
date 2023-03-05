
Public Sub ScpDl()
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    ScpDlParam True
    Call Xftp
    
'    If "On" = ReadRegAR() Then
'        Dim exer As String
'        exer = Cells(currentRow, 16)
'        If InStr(exer, "ScpDl") = 0 Then
'            Cells(currentRow, 16) = Trim(exer & " " & "ScpDl")
'        End If
'    End If
    
    Cells(currentRow, 17) = "Success"
End Sub

