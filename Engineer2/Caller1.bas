
Public Sub Caller1()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler
    Dim n As Integer
    n = Selection.count
    If n > 1 Then
        n = Selection.SpecialCells(xlCellTypeVisible).count
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
    'Dim sProcName As String
    'Dim pk As VBIDE.vbext_ProcKind
    Dim proc As String

    Dim subStr As String

    Dim firstRow As Boolean

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim module As String
    module = Cells(currentRow, 1)
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
        If module = objComponent.Name Then
            'Find the code module for the project.
            Set objCode = objComponent.CodeModule
            'Scan through the code module, looking for procedures.
            iLine = 1
            Do While iLine < objCode.CountOfLines
                codeOfLine = objCode.Lines(iLine, 1)
                proc = objCode.ProcOfLine(iLine, vbext_pk_Proc)

                If subName = proc Then
                    If Not firstRow Then
                        If Trim(codeOfLine) = "" Or True = StartsWith(Trim(codeOfLine), "'") Then
                            firstRow = False
                        Else
                            'MsgBox codeOfLine
                            subStr = SearchRegxKwInStr(codeOfLine, "^P[^ ]+ ([^ ]+)")
                            'MsgBox subStr
                            firstRow = True
                        End If
                    End If
                End If
                iLine = iLine + 1
            Loop
            Set objCode = Nothing
            Set objComponent = Nothing
        End If
    Next objComponent

    Dim funArr() As Variant

    Dim j As Integer
    funArr = CallerSignatures(subName)
    'MsgBox UBound(funArr)
    For j = 1 To UBound(funArr)
        If callerList = "" Then
            callerList = subName & "{" & Left(subStr, 1) & "}" & "<-" & CStr(funArr(j))
        Else
            callerList = callerList & Chr(13) & Chr(10) & subName & "{" & Left(subStr, 1) & "}" & "<-" & CStr(funArr(j))
        End If
    Next j

    'Clean up and exit.
    Set objProject = Nothing

    If callerList <> "" Then
        Cells(currentRow, 8) = callerList & Chr(13) & Chr(10)
    Else
        Cells(currentRow, 8) = ""
    End If

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

