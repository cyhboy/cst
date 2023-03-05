
Public Sub ScpUl()
    If testing Then
        Exit Sub
    End If
    ScpUlParam True
    Call Xftp
    
'    If "On" = ReadRegAR() Then
'        Dim currentRow As Integer
'        currentRow = ActiveCell.Row
'        Dim exer As String
'        exer = Cells(currentRow, 16)
'        If InStr(exer, "ScpUl") = 0 Then
'            Cells(currentRow, 16) = Trim(exer & " " & "ScpUl")
'        End If
'    End If
    
End Sub

