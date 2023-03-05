
Public Function ReadPropertyInAppFiles(filename As String)
    If testing Then
        Exit Function
    End If

    Dim fso, dicProps, oTS As Variant
    Dim sPFSpec, sSect As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    sPFSpec = GetAppDrive() & "\" & filename
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

