
Public Function ReadPropertyInAppFiles(fileName As String)
    If testing Then Exit Function
    Dim fso, sPFSpec, dicProps, oTS, sSect
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    sPFSpec = GetAppDrive() & "\" & fileName
    'MsgBox sPFSpec
    Set dicProps = CreateObject("Scripting.Dictionary")
    Set oTS = fso.OpenTextFile(sPFSpec)
    sSect = ""
    
    Do Until oTS.AtEndOfStream
        Dim sLine: sLine = Trim(oTS.readline)
        If "" <> sLine Then
            If "#" = Left(sLine, 1) Then
                sSect = sLine
            Else
                If "" = sSect Then
                
                Else
                    Dim aParts: aParts = Split(sLine, "=")
                    If 2 = UBound(aParts) Then
                        dicProps(Trim(aParts(0))) = aParts(1) & "=" & aParts(2)
                    ElseIf 1 = UBound(aParts) Then
                        dicProps(Trim(aParts(0))) = aParts(1)
                    Else
                    End If
                End If
            End If
        End If
    Loop
    oTS.Close
    Set fso = Nothing
    Set ReadPropertyInAppFiles = dicProps
End Function
