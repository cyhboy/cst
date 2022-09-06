
Public Sub TouchDoc()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    
    Dim filePath As String
    filePath = Cells(currentRow, 9) & Cells(currentRow, 11)
    
    Dim wa As New Word.Application
    Dim wd As Word.Document
    Dim objSelection As Word.Selection
    
    wa.Visible = False
    
    Set wd = wa.Documents.Open(filePath)
    
'    Dim objComment As Word.comment
'    For Each objComment In wd.Comments
'        MsgBox objComment.Author
'        objComment.Author = ""
'        objComment.Initial = ""
'    Next
    
    Set objSelection = wa.Selection
    
    objSelection.Font.Bold = True
    
    objSelection.Font.Size = "22"
    
    objSelection.TypeText ("I am new here" & vbCrLf)
    
    wd.Save
    'Sleep 3000
    wd.Close
    'wd.Close savechanges:=True
    wa.Quit
    Set wa = Nothing
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 10
    End If
End Sub


