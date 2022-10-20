
Public Sub TagProc()
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
                RobotRunByParam "TagProc"
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
    
    Dim folder As String
    Dim resultStr As String
    folder = "C:\SANDBOX\VB_SPACE\GIT_CST_PROJECT\" & Format(Now, "yyyyMMdd") & "\" & module & "\"
    resultStr = ReadLineByFile(folder & subb & ".bas")
    
    Dim funcStr As String
    Dim subStr As String
    Dim otherStr As String
    
    If MatchRegx(resultStr, "^Public Sub ") Or MatchRegx(resultStr, "^Public Function ") Then
        If MatchRegx(resultStr, "^ *If testing Then") Then
            Cells(currentRow, 17) = "TESTING"
        Else
            Cells(currentRow, 17) = "TESTER"
        End If
    Else
        Cells(currentRow, 17) = "EXEMPT"
    End If
    
    If MatchRegx(resultStr, "^ *Shell ") Then
        Cells(currentRow, 18) = "Shell"
    End If

    If MatchRegx(resultStr, "^ *Set objshell = CreateObject\(""Wscript.Shell""\)") Then
        Cells(currentRow, 18) = Cells(currentRow, 18) & "Wscript.Shell"
    End If
    
    
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

'    If MatchRegx(resultStr, "^ *On Error GoTo ErrorHandler") Then
'        Cells(currentRow, 21) = "ErrorHandler"
'    ElseIf MatchRegx(resultStr, "^ *On Error Resume Next") Then
'        Cells(currentRow, 21) = "ErrorResume"
'    Else
'        Cells(currentRow, 21) = "ErrorUncapture"
'    End If

'    If MatchRegx(resultStr, "^ *On Error Resume Next") Then
'        Cells(currentRow, 22) = "ErrorResume"
'    Else
'        Cells(currentRow, 22) = "ErrorThrow"
'    End If
    Cells(currentRow, 22) = ""
    
'    If MatchRegx(resultStr, "^ *On Error GoTo LineHandler") Then
'        Cells(currentRow, 23) = "SoftCode"
'    Else
'        Cells(currentRow, 23) = "HardCode"
'    End If
    Cells(currentRow, 23) = ""
    
'    If MatchRegx(resultStr, "^ *MsgBox ") Then
'        Cells(currentRow, 28) = "MsgBox Alert"
'    End If
    
    Cells(currentRow, 28) = ""
    
    TagProcRun resultStr, "^ *(On Error .*)", True, True, 21
    
    TagProcRun resultStr, "^ *n = Selection.SpecialCells\(xlCellTypeVisible\)\.count", True, False, 24
    
    TagProcRun resultStr, "^ *(Set .* = CreateObject\(""Scripting.FileSystemObject""\))", True, True, 25
        
    TagProcRun resultStr, "^ *(Set objWMI = GetObject.*)", True, True, 26

    TagProcRun resultStr, "^ *(cn.Open .*)", True, True, 27
    
    TagProcRun resultStr, "^ *[^ ]+ = MyQuestionBox\([^,\r]+\)", True, True, 29
    
    TagProcRun resultStr, "^ *Set fso = Nothing", True, False, 30

    TagProcRun resultStr, "^ *MsgBox ""Please setup repository database. """, True, False, 31
    
    TagProcRun resultStr, "[ \(]ActiveWorkbook.FullName", True, False, 32
        
    TagProcRun resultStr, "[\.]Application.Cells.Find", True, False, 33

    TagProcRun resultStr, "(SearchRegxKwInStrMultToList\([^,\r]+, [^,\r]+, [^,\r]+, [^,\r\)]+\))", True, True, 36
    
    TagProcRun resultStr, "(SearchRegxKwInStr\([^,\r]+, [^,\r]+\))", True, True, 37

    TagProcRun resultStr, "(SearchRegxKwInFileMultToList\([^,\r]+, [^,\r]+, [^,\r\)]+\))", True, True, 38

    TagProcRun resultStr, "(SearchRegxKwInStrToList\([^,\r]+, [^,\r]+\))", True, True, 39
    
    TagProcRun resultStr, "(SearchRegxKwInFile\([^,\r]+, [^,\r]+\))", True, True, 40

    TagProcRun resultStr, "\\([^ ""\.\\]+\.vbs)", True, True, 41
    
    TagProcRun resultStr, "\\([^ ""\.\\]+\.jar)", True, True, 42
     
    TagProcRun resultStr, "\\([^ ""\.\\]+\.exe)", True, True, 43
    
    TagProcRun resultStr, "\\([^ ""\.\\]+\.ps1)", True, True, 44
    
    TagProcRun resultStr, "^ *(Set .* = CreateObject\(""Shell.Application""\).*)", True, False, 45

    TagProcRun resultStr, "^ *(Set .* = CreateObject\(""InternetExplorer.Application""\).*)", True, False, 46
    
ErrorHandler:
    If Err.Number <> 0 Then
        Cells(currentRow, 47) = Err.Description
    End If
End Sub

