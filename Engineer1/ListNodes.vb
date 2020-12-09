
Public Function ListNodes(node As IXMLDOMNode, procName As String, iCol As Integer, display As Boolean)
    'On Error GoTo ErrorHandler
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim attr As IXMLDOMAttribute
    
    If node.nodeName = "#comment" Then
        Exit Function
    End If
    
    Set attr = node.Attributes.getNamedItem("label")
    'MsgBox node.nodeName
    If (Not attr Is Nothing) Then
        If iCol = 6 Or iCol = 7 Then
            If node.nodeName <> "mso:menu" Then
                buttonNum = buttonNum + 1
                If testing Then
                    If attr.text <> "Ver" And attr.text <> "Test" And attr.text <> "TestVBA" Then
                        'MsgBox attr.text
                        TestCall attr.text
                    End If
                End If
            Else
                menuNum = menuNum + 1
            End If
        End If
        
        If display Then
            Cells(currentRow, iCol) = attr.text
        End If
        
        If attr.text = procName Then
            ListNodes = "exit"
            Exit Function
        End If
    End If
    
    If iCol = 4 Then
        tabNum = tabNum + 1
    End If
    
    If iCol = 5 Then
        groupNum = groupNum + 1
    End If
    
    If node.HasChildNodes() Then
       'MsgBox node.nodeName & " has child nodes"
       Dim node2 As IXMLDOMNode
       For Each node2 In node.ChildNodes
          ListNodes = ListNodes(node2, procName, iCol + 1, display)
          If (ListNodes = "exit") Then
            Exit For
          Else
            If iCol = 6 Then
                Cells(currentRow, iCol + 1) = ""
            End If
          End If
       Next
       'MsgBox "Done listing child nodes for " & node.nodeName
    End If
    
'ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Function

