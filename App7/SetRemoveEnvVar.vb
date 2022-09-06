
Public Sub SetRemoveEnvVar()
    If testing Then
        Exit Sub
    End If
    Dim nameParam As String, valueParam As String, userParam As String
    Dim currentRow As Integer

    currentRow = ActiveCell.Row

    nameParam = Cells(currentRow, 9)
    valueParam = Cells(currentRow, 10)

    userParam = Cells(currentRow, 1)

    Dim objWMI As Object
    Dim objVar As Object

    Set objWMI = GetObject("winmgmts://./root/cimv2:Win32_Environment")
    Set objVar = objWMI.SpawnInstance_
    objVar.Name = nameParam
    objVar.VariableValue = valueParam
    objVar.UserName = userParam
    'objVar.SystemVariable      = False
    'objVar.Caption      = "GUANGZHOU\asnphpb\JAVA_HOME"
    'objVar.Description      = "GUANGZHOU\asnphpb\JAVA_HOME"
    objVar.Put_

    Set objVar = Nothing
    Set objWMI = Nothing
End Sub

