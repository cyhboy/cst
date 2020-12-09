
Public Sub TestCall(proc As String)
    On Error GoTo ErrorHandler
    Application.Run "'cst.xlam'!" & proc
    
ErrorHandler:
    If Err.Number <> 0 Then
        Err.Clear
        'Application.Run "'cst.xlam'!AIA." & proc
        
    Else
        Exit Sub
    End If
    

    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

