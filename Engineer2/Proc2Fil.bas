
Public Sub Proc2Fil()
    If testing Then
        Exit Sub
    End If

    On Error GoTo ErrorHandler
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
                RobotRunByParam "Proc2Fil"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim module As String
    module = Cells(currentRow, 1)
    Dim subb As String
    subb = Cells(currentRow, 2)

    Dim i As Long
    Dim proc As String
    Dim VBComp As VBComponent
    Dim objProject As VBIDE.VBProject
    Dim objCode As VBIDE.CodeModule

    Dim codeOfLine As String
    'Dim startOfProc As Long
    'Dim lengthOfProc As Long
    'startOfProc = objCode.ProcStartLine(proc, vbext_pk_Proc)
    'lengthOfProc = objCode.ProcCountLines(proc, vbext_pk_Proc)
    Dim resultStr As String
    Set objProject = ThisWorkbook.VBProject
    For Each VBComp In objProject.VBComponents
        If module = VBComp.Name Then
            ' Find the code module for the project.
            Set objCode = VBComp.CodeModule
            For i = 1 To objCode.CountOfLines
                codeOfLine = objCode.Lines(i, 1)
                'If Trim(codeOfLine) <> "" Then
                proc = objCode.ProcOfLine(i, vbext_pk_Proc)
                If subb = proc Then
                    resultStr = resultStr & codeOfLine & Chr(13) & Chr(10)
                End If
                'End If
            Next i
        End If
    Next VBComp

     Dim folder As String

    'folder = "C:\SANDBOX\VB_SPACE\CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"
    'folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"
    folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"

    CreateFolder folder

    WriteTxt2Tmp resultStr, folder & subb & ".bas"
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 5
        Cells(currentRow, 47) = "###"
    End If
End Sub

