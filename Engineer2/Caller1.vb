
Public Sub Caller1()
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
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
                RobotRunByParam "Caller1"
            End If
        Next curCell
        Exit Sub
    End If
    'Declare other miscellaneous variables.
    Dim iLine As Integer
    Dim sProcName As String
    Dim pk As VBIDE.vbext_ProcKind
    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim subName As String
    subName = Cells(currentRow, 2)
    Dim callerList As String
    callerList = ""
    Dim objProject As VBIDE.VBProject
    Set objProject = ThisWorkbook.VBProject
    'Iterate through each component in the project.
    Dim objComponent As VBIDE.VBComponent
    Dim objCode As VBIDE.CodeModule
    Dim codeOfLine As String
    For Each objComponent In objProject.VBComponents
        'Find the code module for the project.
        Set objCode = objComponent.CodeModule
        'Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < objCode.CountOfLines
            codeOfLine = objCode.Lines(iLine, 1)
            If Trim(codeOfLine) <> "" And False = StartsWith(Trim(codeOfLine), "'") Then
                If Trim(codeOfLine) = "Call " & subName _
                Or (InStr(Trim(codeOfLine), subName & " ") = 1 And InStr(Trim(codeOfLine), " = ") = 0) _
                Or InStr(codeOfLine, " = " & subName & "(") > 0 _
                Or InStr(codeOfLine, " <> " & subName & "(") > 0 _
                Or InStr(codeOfLine, " & " & subName & "(") > 0 _
                Or InStr(codeOfLine, "(" & subName & "(") > 0 _
                Or InStr(codeOfLine, " - " & subName & "(") > 0 _
                Or InStr(codeOfLine, ", " & subName & "(") > 0 Then
                    sProcName = objCode.ProcOfLine(iLine, pk)
                    If callerList = "" Then
                        callerList = subName & "->" & sProcName
                    Else
                        callerList = callerList & Chr(13) & Chr(10) & subName & "->" & sProcName
                    End If
                End If
            End If
            iLine = iLine + 1
        Loop
        Set objCode = Nothing
        Set objComponent = Nothing
    Next
    'Clean up and exit.
    Set objProject = Nothing
    Cells(currentRow, 8) = callerList
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

