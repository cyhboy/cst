
Public Function GetWorkbookProperties(ByVal filePath As String, ByVal propName As String)
    If testing Then
        Exit Function
    End If
    Dim retvalue As String
    Dim appOffice As New Application
    Dim richFile As Workbook
    Set richFile = appOffice.Workbooks.Open(filePath)
    retvalue = richFile.BuiltinDocumentProperties(propName)
    richFile.Saved = True
    'richFile.Close
    appOffice.Workbooks.Close
    appOffice.Quit
    Set appOffice = Nothing
    GetWorkbookProperties = retvalue
End Function

