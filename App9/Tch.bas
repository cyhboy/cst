
Public Sub Tch()
    If testing Then
        Exit Sub
    End If
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim fileName As String
    fileName = Cells(currentRow, 11)
    Dim localFolder As String
    localFolder = Cells(currentRow, 9)
    Dim txtValue As String
    txtValue = Cells(currentRow, 10)
    Dim fso As Object
    ' Dim txtFile As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath1 As String
    filePath1 = "C:\AppFiles\linux.txt"
    Dim result As String
    If Not fso.fileexists(localFolder & fileName) Then
        result = fso.copyfile(filePath1, localFolder & fileName)
        If result = "" Then
            MsgBox "copied linux file done"
        End If
    Else
        MsgBox "target file is existing, copy canceled"
    End If
    Set fso = Nothing
End Sub

