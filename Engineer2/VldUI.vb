
Public Sub VldUI()
    If testing Then Exit Sub
    'On Error GoTo ErrorHandler
    Dim n As Integer
    n = Selection.Count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).Count
    End If
    If n > 1 Then
        Dim curCell As Range
        For Each curCell In Selection
            If curCell.EntireColumn.Hidden = False And curCell.EntireRow.Hidden = False Then
                curCell.Select
                'MsgBox subName
                RobotRunByParam "VldUI"
            End If
        Next curCell
        Exit Sub
    End If
    
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim procName As String

    procName = Cells(currentRow, 2)

    
    Dim XCMFILE As Variant

    Dim buttonActionName As String

    Dim nod As IXMLDOMNode

    
    Set XCMFILE = CreateObject("Microsoft.XMLDOM")
    
    XCMFILE.Load ("C:\Users\" & Environ$("username") & "\AppData\Local\Microsoft\Office\Excel.officeUI") 'Load XCM File
    
    For Each nod In XCMFILE.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs/mso:tab/mso:group/mso:button")
        buttonActionName = nod.Attributes.getNamedItem("onAction").text 'Search for id attribute within node
        If EndsWith(buttonActionName, "!" & procName) Then
            Cells(currentRow, 6) = nod.Attributes.getNamedItem("label").text
            Cells(currentRow, 5) = nod.ParentNode.Attributes.getNamedItem("label").text
            Cells(currentRow, 4) = nod.ParentNode.ParentNode.Attributes.getNamedItem("label").text
            Set XCMFILE = Nothing
            Exit Sub
        Else
            Cells(currentRow, 6) = "N/A"
            Cells(currentRow, 5) = "N/A"
            Cells(currentRow, 4) = "N/A"
        End If
    Next
    
    For Each nod In XCMFILE.SelectNodes("//mso:customUI/mso:ribbon/mso:tabs/mso:tab/mso:group/mso:menu/mso:button")
        buttonActionName = nod.Attributes.getNamedItem("onAction").text 'Search for id attribute within node
        If EndsWith(buttonActionName, "!" & procName) Then
            Cells(currentRow, 7) = nod.Attributes.getNamedItem("label").text
            Cells(currentRow, 6) = nod.ParentNode.Attributes.getNamedItem("label").text
            Cells(currentRow, 5) = nod.ParentNode.ParentNode.Attributes.getNamedItem("label").text
            Cells(currentRow, 4) = nod.ParentNode.ParentNode.ParentNode.Attributes.getNamedItem("label").text
            Set XCMFILE = Nothing
            Exit Sub
        Else
            Cells(currentRow, 7) = "N/A"
            Cells(currentRow, 6) = "N/A"
            Cells(currentRow, 5) = "N/A"
            Cells(currentRow, 4) = "N/A"
        End If
    Next
    

End Sub

