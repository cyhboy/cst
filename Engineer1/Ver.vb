
Public Sub Ver()
    testing = True
    'On Error GoTo ErrorHandler
    Call CntOfficeUI
    Dim fso, fileObject As Object
    Dim addInsPath As String
'    addInsPath = "C:\Program Files\Microsoft Office\Office14\Library\cst.xlam"
'    If Is64bit Then
'        addInsPath = "C:\Program Files (x86)\Microsoft Office\Office14\Library\cst.xlam"
'    End If
    addInsPath = "C:\AppFiles\cst.xlam"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileObject = fso.GetFile(addInsPath)
    Dim Info As String
    Info = "Thanks for choosing Common Support Toolkits! " & vbCrLf
    Info = Info & "The release timestamp of your current copy is " & fileObject.DateLastModified & ". " & vbCrLf
    Info = Info & "Current total number of tab definition is " & tabNum & ". " & vbCrLf
    Info = Info & "Current total number of group definition is " & groupNum & ". " & vbCrLf
    Info = Info & "Current total number of button definition is " & buttonNum & ". " & vbCrLf
    Info = Info & "Current total number of menu definition is " & menuNum & ". " & vbCrLf
    
    MsgBox Info, vbInformation, "Version"
    Set fileObject = Nothing
    Set fso = Nothing
'ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
    testing = False
End Sub

