
Public Sub LookupIp()
    If testing Then
        Exit Sub
    End If

    Dim parameter As String
    Dim currentRow As Integer

    currentRow = ActiveCell.Row
    parameter = Cells(currentRow, 2)

    Dim objshell, objExec As Variant
    Dim strCmd, strLine, strIP, strFQDN As String
    Set objshell = CreateObject("Wscript.Shell")

    strCmd = "nslookup " & parameter & """"
    Set objExec = objshell.Exec(strCmd)


    Do Until objExec.StdOut.AtEndOfStream
        strLine = objExec.StdOut.readline()
        If (Left(strLine, 8) = "Address:") Then
            strIP = Trim(Mid(strLine, 9))
        End If
    Loop

    If Cells(currentRow, 6).Value = "" Or Cells(currentRow, 6).Value = strIP Then
        Cells(currentRow, 6).Value = strIP
    Else
        Cells(currentRow, 6).Value = Cells(currentRow, 6).Value & Chr(10) & strIP
    End If
End Sub


