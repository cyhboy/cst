
Public Function URLEncode2(StringToEncode As String, Optional UsePlusRatherThanHexForSpace As Boolean = False) As String
    If testing Then
        Exit Function
    End If

    Dim TempAns As String
    Dim CurChr As Integer
    CurChr = 1

    Do Until CurChr - 1 = Len(StringToEncode)
        Select Case Asc(Mid(StringToEncode, CurChr, 1))
         Case 48 To 57, 65 To 90, 97 To 122
            TempAns = TempAns & Mid(StringToEncode, CurChr, 1)
         Case 32
            If UsePlusRatherThanHexForSpace = True Then
                TempAns = TempAns & "+"
            Else
                TempAns = TempAns & "%" & Hex(32)
            End If
         Case Else
            TempAns = TempAns & "%" & _
            Right("0" & Hex(Asc(Mid(StringToEncode, _
            CurChr, 1))), 2)
        End Select

        CurChr = CurChr + 1
    Loop

    URLEncode2 = TempAns
End Function

