
Public Function GetPingResult(host As String)
    If testing Then
        Exit Function
    End If

    Dim objWMI As Object
    Dim objStatus As Object
    Dim result As String
    Dim strResult As String

    Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("Select * from Win32_PingStatus Where Address = '" & host & "'")

    For Each objStatus In objWMI
        Select Case objStatus.StatusCode
         Case 0: strResult = "Connected"
         Case 11001: strResult = "Buffer too small"
         Case 11002: strResult = "Destination net unreachable"
         Case 11003: strResult = "Destination host unreachable"
         Case 11004: strResult = "Destination protocol unreachable"
         Case 11005: strResult = "Destination port unreachable"
         Case 11006: strResult = "No resources"
         Case 11007: strResult = "Bad option"
         Case 11008: strResult = "Hardware error"
         Case 11009: strResult = "Packet too big"
         Case 11010: strResult = "Request timed out"
         Case 11011: strResult = "Bad request"
         Case 11012: strResult = "Bad route"
         Case 11013: strResult = "Time-To-Live (TTL) expired transit"
         Case 11014: strResult = "Time-To-Live (TTL) expired reassembly"
         Case 11015: strResult = "Parameter problem"
         Case 11016: strResult = "Source quench"
         Case 11017: strResult = "Option too big"
         Case 11018: strResult = "Bad destination"
         Case 11032: strResult = "Negotiating IPSEC"
         Case 11050: strResult = "General failure"
         Case Else: strResult = "Unknown host"
        End Select
        GetPingResult = strResult
    Next objStatus

    Set objWMI = Nothing
End Function

