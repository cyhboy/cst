
Public Function CallerSignatures(procName As String) As Variant
    If testing Then
        Exit Function
    End If
    Dim objProject As VBIDE.VBProject
    Set objProject = ThisWorkbook.VBProject

    Dim objComponent As VBIDE.VBComponent
    Dim objCode As VBIDE.CodeModule

    Dim iLine As Integer
    Dim codeOfLine As String

    Dim funArr() As Variant
    ReDim funArr(0)

    Dim proc As String

    Dim proccStr As String

    Dim firstRow As Boolean

    Dim codeRow As Integer

    Dim funCount As Integer
    funCount = 0

    For Each objComponent In objProject.VBComponents
        'Find the code module for the project.
        Set objCode = objComponent.CodeModule
        'Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < objCode.CountOfLines
            codeOfLine = objCode.Lines(iLine, 1)
            If proc <> objCode.ProcOfLine(iLine, vbext_pk_Proc) Then
                proc = objCode.ProcOfLine(iLine, vbext_pk_Proc)
                firstRow = False
                codeRow = 0
            End If

            'If module = objComponent.Name Then

            If Not firstRow Then
                If Trim(codeOfLine) = "" Or True = StartsWith(Trim(codeOfLine), "'") Then
                    firstRow = False
                Else
                    proccStr = SearchRegxKwInStr(codeOfLine, "^P[^ ]+ ([^ ]+)")
                    firstRow = True
                End If
            End If

            'End If

            If Trim(codeOfLine) <> "" And False = StartsWith(Trim(codeOfLine), "'") Then
                If Trim(codeOfLine) = "Call " & procName _
                    Or (InStr(Trim(codeOfLine), procName & " ") = 1 And InStr(Trim(codeOfLine), " = ") = 0) _
                    Or InStr(codeOfLine, " = " & procName & "(") > 0 _
                    Or InStr(codeOfLine, " <> " & procName & "(") > 0 _
                    Or InStr(codeOfLine, " & " & procName & "(") > 0 _
                    Or InStr(codeOfLine, "(" & procName & "(") > 0 _
                    Or InStr(codeOfLine, " - " & procName & "(") > 0 _
                    Or InStr(codeOfLine, ", " & procName & "(") > 0 _
                    Or InStr(codeOfLine, "If " & procName & "(") > 0 Then
                    'sProcName = objCode.ProcOfLine(iLine, pk)
                    funCount = funCount + 1
                    ReDim Preserve funArr(0 To funCount)
                    'MsgBox proc & "{" & Left(proccStr, 1) & "}" & "(" & codeRow & ")"
                    funArr(funCount) = proc & "{" & Left(proccStr, 1) & "}" & "(" & codeRow & ")"

                End If
            End If
            codeRow = codeRow + 1
            iLine = iLine + 1
        Loop
        Set objCode = Nothing
        Set objComponent = Nothing
    Next objComponent
    CallerSignatures = funArr
End Function

