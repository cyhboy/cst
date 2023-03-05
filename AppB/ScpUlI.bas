
Public Sub ScpUlI()
    If testing Then
        Exit Sub
    End If
    ScpUlParam True
    Call XftpI
    
'    If "On" = ReadRegAR() Then
'        Dim currentRow As Integer
'        currentRow = ActiveCell.Row
'        Dim exer As String
'        exer = Cells(currentRow, 16)
'        If InStr(exer, "ScpUlI") = 0 Then
'            Cells(currentRow, 16) = Trim(exer & " " & "ScpUlI")
'        End If
'    End If
    
End Sub

