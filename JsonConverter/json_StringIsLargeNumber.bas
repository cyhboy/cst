
Private Function json_StringIsLargeNumber(json_String As Variant) As Boolean
    If testing Then Exit Function
    ' Check if the given string is considered a "large number"
    ' (See json_ParseNumber)

    Dim json_Length As Long
    Dim json_CharIndex As Long
    json_Length = VBA.Len(json_String)

    ' Length with be at least 16 characters and assume will be less than 100 characters
    If json_Length >= 16 And json_Length <= 100 Then
        Dim json_CharCode As String

        json_StringIsLargeNumber = True

        For json_CharIndex = 1 To json_Length
            json_CharCode = VBA.Asc(VBA.Mid$(json_String, json_CharIndex, 1))
            Select Case json_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                json_StringIsLargeNumber = False
                Exit Function
            End Select
        Next json_CharIndex
    End If
End Function

