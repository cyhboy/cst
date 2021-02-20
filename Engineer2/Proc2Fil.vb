
Public Sub Proc2Fil()

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
    Next
    Dim funcStr As String
    Dim subStr As String
    Dim otherStr As String
    If MatchRegx(resultStr, "^P.* Function [^\(]+\(") Then
        
        funcStr = SearchRegxKwInStr(resultStr, "^(P[^ ]+ Function [^\(]+\(.*\).*)", True)
        Cells(currentRow, 19) = funcStr
        If InStr(funcStr, "()") > 0 Then
            Cells(currentRow, 20) = 0
        Else
            Cells(currentRow, 20) = CntSubstring(funcStr, ", ") + 1
        End If
    End If
    
    If MatchRegx(resultStr, "^P.* Sub [^\(]+\(") Then

        subStr = SearchRegxKwInStr(resultStr, "^(P[^ ]+ Sub [^\(]+\(.*\).*)", True)
        Cells(currentRow, 19) = subStr
        If InStr(subStr, "()") > 0 Then
            Cells(currentRow, 20) = 0
        Else
            Cells(currentRow, 20) = CntSubstring(subStr, ", ") + 1
        End If
    End If
    
    If MatchRegx(resultStr, "^P.* Property Get [^\(]+\(") Then
        
        otherStr = SearchRegxKwInStr(resultStr, "^(P[^ ]+ Property Get [^\(]+\(.*\).*)", True)
        Cells(currentRow, 19) = otherStr
        If InStr(otherStr, "()") > 0 Then
            Cells(currentRow, 20) = 0
        Else
            Cells(currentRow, 20) = CntSubstring(otherStr, ", ") + 1
        End If
    End If
    
    If MatchRegx(resultStr, "^Public Sub ") Or MatchRegx(resultStr, "^Public Function ") Then
        If MatchRegx(resultStr, "^ *If testing Then Exit") Then
            Cells(currentRow, 17) = "TESTING"
        Else
            Cells(currentRow, 17) = "TESTER"
        End If
    Else
        Cells(currentRow, 17) = "EXEMPT"
    End If
    
    
    If MatchRegx(resultStr, "^ *cn.Open ") Then

        Cells(currentRow, 27) = Join(SearchRegxKwInStrMultToList(resultStr, "^ *(cn.Open .*)", 0, True), Chr(10))
    End If
    
'    If InStr(resultStr, "Array(""staffid"")") > 0 Then
'        Cells(currentRow, 18) = "staffid"
'    End If
'    If InStr(resultStr, "Array(""alias"")") > 0 Then
'        Cells(currentRow, 18) = "alias"
'    End If
'    If InStr(resultStr, "Array(""sectionID"")") > 0 Then
'        Cells(currentRow, 18) = "sectionID"
'    End If
'    If InStr(resultStr, "Array(""terminalID"")") > 0 Then
'        Cells(currentRow, 18) = "terminalID"
'    End If
'    If InStr(resultStr, "Array(""computerName"")") > 0 Then
'        Cells(currentRow, 18) = "computerName"
'    End If

    
    If MatchRegx(resultStr, "^ *Shell ") Then
        Cells(currentRow, 18) = "Shell"
    End If
    
    If MatchRegx(resultStr, "^ *Set objshell = CreateObject\(""Wscript.Shell""\)") Then
        Cells(currentRow, 18) = Cells(currentRow, 18) & "Wscript.Shell"
    End If
    
    If MatchRegx(resultStr, "^ *Set .* = CreateObject\(""Scripting.FileSystemObject""\)") Then
        Cells(currentRow, 25) = SearchRegxKwInStr(resultStr, "^ *(Set .* = CreateObject\(""Scripting.FileSystemObject""\))", True)
    End If
    
    

    
    If MatchRegx(resultStr, "^ *Set fso = Nothing") Then
        Cells(currentRow, 30) = "Set fso = Nothing"
    End If
    
    If MatchRegx(resultStr, "^ *Set objWMI = GetObject") Then
        Cells(currentRow, 26) = SearchRegxKwInStr(resultStr, "^ *(Set objWMI = GetObject.*)", True)
    End If
    
    If MatchRegx(resultStr, "^ *MsgBox ") Then
        Cells(currentRow, 28) = "MsgBox Alert"
    End If
    
    If MatchRegx(resultStr, "^ *[^ ]+ = MsgBox\(") Then
        Cells(currentRow, 29) = "MsgBox Question"
    End If
    
    If MatchRegx(resultStr, "^ *On Error GoTo ErrorHandler") Then
        Cells(currentRow, 21) = "ErrorHandler"
    Else
        Cells(currentRow, 21) = "ErrorUncapture"
    End If
    
    If MatchRegx(resultStr, "^ *On Error Resume Next") Then
        Cells(currentRow, 22) = "ErrorResume"
    Else
        Cells(currentRow, 22) = "ErrorThrow"
    End If
    
    If MatchRegx(resultStr, "^ *On Error GoTo LineHandler") Then
        Cells(currentRow, 23) = "SoftCode"
    Else
        Cells(currentRow, 23) = "HardCode"
    End If
    
    If MatchRegx(resultStr, "^ *MsgBox ""Please setup repository database. """) Then
        Cells(currentRow, 31) = "Require Usage Record"
    End If
    
   
    If MatchRegx(resultStr, "[ \(]ActiveWorkbook.FullName") Then
        Cells(currentRow, 32) = "ActiveWorkbook.FullName"
    End If
    
    If MatchRegx(resultStr, "[\.]Application.Cells.Find") Then
        Cells(currentRow, 33) = "Application.Cells.Find"
    End If

    
    If MatchRegx(resultStr, "SearchRegxKwInStrMultToList\(") Then
        Cells(currentRow, 36) = Join(SearchRegxKwInStrMultToList(resultStr, "(SearchRegxKwInStrMultToList\([^,\r]+, [^,\r]+, [^,\r]+, [^,\r\)]+\))", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "SearchRegxKwInStr\(") Then
        Cells(currentRow, 37) = Join(SearchRegxKwInStrMultToList(resultStr, "(SearchRegxKwInStr\([^,\r]+, [^,\r]+\))", 0, False), Chr(10))
    End If
    
    If MatchRegx(resultStr, "SearchRegxKwInFileMultToList\(") Then
        Cells(currentRow, 38) = Join(SearchRegxKwInStrMultToList(resultStr, "(SearchRegxKwInFileMultToList\([^,\r]+, [^,\r]+, [^,\r\)]+\))", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "SearchRegxKwInStrToList\(") Then
        Cells(currentRow, 39) = Join(SearchRegxKwInStrMultToList(resultStr, "(SearchRegxKwInStrToList\([^,\r]+, [^,\r]+\))", 0, False), Chr(10))
    End If
    
    If MatchRegx(resultStr, "SearchRegxKwInFile\(") Then
        Cells(currentRow, 40) = Join(SearchRegxKwInStrMultToList(resultStr, "(SearchRegxKwInFile\([^,\r]+, [^,\r]+\))", 0, True), Chr(10))
    End If

    If MatchRegx(resultStr, "\\[^ ""\.\\]+\.vbs") Then
        Cells(currentRow, 41) = Join(SearchRegxKwInStrMultToList(resultStr, "\\([^ ""\.\\]+\.vbs)", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "\\[^ ""\.\\]+\.jar") Then
        Cells(currentRow, 42) = Join(SearchRegxKwInStrMultToList(resultStr, "\\([^ ""\.\\]+\.jar)", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "\\[^ ""\.\\]+\.exe") Then
        Cells(currentRow, 43) = Join(SearchRegxKwInStrMultToList(resultStr, "\\([^ ""\.\\]+\.exe)", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "\\[^ ""\.\\]+\.ps1") Then
        Cells(currentRow, 44) = Join(SearchRegxKwInStrMultToList(resultStr, "\\([^ ""\.\\]+\.ps1)", 0, True), Chr(10))
    End If
    
    If MatchRegx(resultStr, "^ *Set .* = CreateObject\(""Shell.Application""\).*") Then
        Cells(currentRow, 45) = SearchRegxKwInStr(resultStr, "^ *(Set .* = CreateObject\(""Shell.Application""\).*)", True)
    End If
    
    If MatchRegx(resultStr, "^ *Set .* = CreateObject\(""InternetExplorer.Application""\).*") Then
        Cells(currentRow, 46) = SearchRegxKwInStr(resultStr, "^ *(Set .* = CreateObject\(""InternetExplorer.Application""\).*)", True)
    End If
    
    
    If MatchRegx(resultStr, "^ *n = Selection.SpecialCells(xlCellTypeVisible).count") Then
        Cells(currentRow, 24) = "Rundown"
    Else
        Cells(currentRow, 24) = "Singleton"
    End If
    
    Dim folder As String
    
    'folder = "C:\SANDBOX\VB_SPACE\CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"
    folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"
    
    CreateFolder folder

    WriteTxt2Tmp resultStr, folder & subb & ".vb"
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 5
        Cells(currentRow, 47) = "###"
    End If
End Sub

