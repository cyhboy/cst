
Public Sub Ribbon()
    If testing Then
        Exit Sub
    End If
    Call UnHF
    'Call Frz
    Call CleanRgn
    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler
    'Dim VBAEditor As VBIDE.VBE
    Dim objProject As VBIDE.VBProject
    Dim objComponent As VBIDE.VBComponent
    Dim objCode As VBIDE.CodeModule

    ' Declare other miscellaneous variables.
    Dim iLine As Integer
    Dim sProcName As String
    Dim pk As VBIDE.vbext_ProcKind

    Dim i As Integer
    i = 2
    'Set VBAEditor = Application.VBE

    Dim codeOfLine As String
    'Get the project details in the workbook.
    Set objProject = ThisWorkbook.VBProject

    'Iterate through each component in the project.
    For Each objComponent In objProject.VBComponents
        'Find the code module for the project.
        Set objCode = objComponent.CodeModule
        'Scan through the code module, looking for procedures.
        iLine = 1
        Do While iLine < objCode.CountOfLines
            codeOfLine = objCode.Lines(iLine, 1)
            If Trim(codeOfLine) <> "" Then
                sProcName = objCode.ProcOfLine(iLine, pk)
                If sProcName <> "" Then
                    'MsgBox objComponent.Name & ": " & sProcName
                    ActiveSheet.Cells(i, 1).Value = objComponent.Name
                    ActiveSheet.Cells(i, 2).Value = sProcName

                    ActiveSheet.Cells(i, 3).Value = iLine

                    iLine = iLine + objCode.ProcCountLines(sProcName, pk) - 2

                    ActiveSheet.Cells(i, 11).Value = iLine
                    'ActiveSheet.Cells(i, 13).FormulaR1C1 = "=RC[-2] - RC[-10] + 1"

                    ActiveSheet.Cells(i, 12).FormulaR1C1 = "=TODAY() + 30 * 2"

                    ActiveSheet.Cells(i, 14).FormulaR1C1 = "=RIGHT(CELL(""filename"", R1C1),LEN(CELL(""filename"", R1C1))-FIND(""]"",CELL(""filename"", R1C1)))"

                    ActiveSheet.Cells(i, 15) = "'" & LPad((i - 1) & "", 4, "0")

                    ActiveSheet.Cells(i, 16) = "#Proc2Fil TagProc #Proc2Md"
                    ActiveSheet.Cells(i, 17) = "ONGOING"

                    i = i + 1
                End If
            End If
            iLine = iLine + 1
        Loop
        Set objCode = Nothing
        Set objComponent = Nothing
    Next objComponent
    'Clean up and exit.
    Set objProject = Nothing

ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If

    Application.ScreenUpdating = True
End Sub

