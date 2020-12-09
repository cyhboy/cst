
Public Function CommonGetTheDrive()
    If testing Then Exit Function
    Dim fso As Object
    Dim obj As Object
    Dim retDrive As String
    retDrive = ""

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    For Each obj In fso.Drives
        'If obj.DriveType = 3 Then
        If obj.path <> "C:" Then
            'MsgBox "Testing Drive: " & obj.path
            'If Dir(obj.path & "\AppFiles\SupportSetup\cst.xlam") <> "" Then
            If fso.fileexists(obj.path & "\AppFiles\SupportSetup\cst.xlam") Then
                retDrive = obj.path
                Exit For
                If obj.path = "M:" Or obj.path = "m:" Then
                    Exit For
                End If
            End If
        End If
    Next
    Set fso = Nothing
    
    'If retDrive = "" Then
    '    retDrive = "\\10.15.76.73\common_oa"
    'End If
    'MsgBox retDrive
    CommonGetTheDrive = retDrive
End Function

