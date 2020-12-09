
Public Sub CntOfficeUI()
    'On Error GoTo ErrorHandler
    Dim procName As String
    procName = "N/A"
    
    Dim iCol As Integer
    iCol = 3
    
    Dim oXML As Object
    Set oXML = New DOMDocument
    'Set oXML = New DOMDocument60
    Dim strURL As String, resultStr As String
    
    Dim strFilePath As String
    strFilePath = "C:\Users\" & Environ$("username") & "\AppData\Local\Microsoft\Office\Excel.officeUI"
    
    
    'MsgBox strFilePath
    
    oXML.Load (strFilePath)
    
    Dim oXmlNodes As IXMLDOMNodeList
    
    'Set oXmlNodes = oXML.SelectNodes("//customUI/ribbon/tabs")
    Set oXmlNodes = oXML.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs")
    
    'MsgBox oXmlNodes.Length

    Dim node As IXMLDOMNode
    Dim x As String
    
    buttonNum = 0
    groupNum = 0
    tabNum = 0
    menuNum = 0
    
    For Each node In oXmlNodes
        
        x = ListNodes(node, procName, iCol, False)
        
        If x = "exit" Then
            'MsgBox "exit"
            Exit Sub
        End If
    Next

    Set oXML = Nothing
    
    'MsgBox tabNum
    'MsgBox groupNum
    'MsgBox buttonNum
'ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

