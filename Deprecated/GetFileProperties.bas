
'Public Sub VsUI()
'    'to be legacy
'    If testing Then Exit Sub
'    'On Error GoTo ErrorHandler
'    Dim n As Integer
'    n = Selection.count
'    If n > 1 Then
'        n = Selection.SpecialCells(xlCellTypeVisible).count
'    End If
'    If n > 1 Then
'        Dim curCell As Range
'        For Each curCell In Selection
'            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
'                curCell.Select
'                'MsgBox subName
'                RobotRunByParam "VsUI"
'            End If
'        Next curCell
'        Exit Sub
'    End If
'
'    Application.ScreenUpdating = False
'
'    Dim currentRow As Integer
'    currentRow = ActiveCell.Row
'
'    Dim procName As String
'
'    procName = Cells(currentRow, 2)
'
'    Dim iCol As Integer
'    iCol = 3
'
'    Dim oXML As Object
'    Set oXML = New DOMDocument
'    Dim strURL As String, resultStr As String
'
'    Dim strFilePath As String
'    strFilePath = "C:\Users\" & Environ$("username") & "\AppData\Local\Microsoft\Office\Excel.officeUI"
'
'    oXML.Load (strFilePath)
'
'    Dim oXmlNodes As IXMLDOMNodeList
'
'    Set oXmlNodes = oXML.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs")
'
'    Dim node As IXMLDOMNode
'    Dim x As String
'    buttonNum = 0
'    groupNum = 0
'    tabNum = 0
'    menuNum = 0
'    For Each node In oXmlNodes
'        x = ListNodes(node, procName, iCol, True)
'        'x = ListNodes(node, procName, iCol, False)
'        If x = "exit" Then
'            'MsgBox "exit"
'            Exit Sub
'        End If
'    Next
'    Set oXML = Nothing
'    Cells(currentRow, 4) = "N/A"
'    Cells(currentRow, 5) = "N/A"
'    Cells(currentRow, 6) = "N/A"
'    Cells(currentRow, 7) = "N/A"
'    'MsgBox "done"
''ErrorHandler:
'    If Err.Number <> 0 Then
'        MyMsgBox Err.Number & " " & Err.Description, 30
'    End If
'    Application.ScreenUpdating = True
'End Sub

Public Sub GetFileProperties()
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim videoPath As String
    Dim videoFileName As String
    Dim videoFullFilename As String
    videoPath = Cells(currentRow, 9)
    videoFileName = Cells(currentRow, 11)
    videoFullFilename = videoPath & videoFileName
    Dim resultStr As String
    Dim oFolder, ofPName As Variant
    Dim i As Integer
    With CreateObject("Shell.Application")
        Set oFolder = .Namespace(Left(videoFullFilename, InStrRev(videoFullFilename, "\") - 1))
        Set ofPName = oFolder.ParseName(Right(videoFullFilename, Len(videoFullFilename) - InStrRev(videoFullFilename, "\")))
        For i = 1 To 100
            resultStr = resultStr & Chr(13) & Chr(10) & oFolder.GetDetailsOf(ofPName, i)
        Next i
    End With

    Cells(currentRow, 18) = resultStr
End Sub

