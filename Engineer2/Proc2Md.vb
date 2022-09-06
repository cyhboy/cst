
Public Sub Proc2Md()
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
                RobotRunByParam "Proc2Md"
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

    Dim call1 As String
    call1 = Cells(currentRow, 8)

    Dim call2 As String
    call2 = Cells(currentRow, 9)

    Dim call3 As String
    call3 = Cells(currentRow, 10)

    Dim beCall As String
    beCall = Cells(currentRow, 13)

    Dim menu1 As String
    menu1 = Cells(currentRow, 4)

    Dim menu2 As String
    menu2 = Cells(currentRow, 5)

    Dim menu3 As String
    menu3 = Cells(currentRow, 6)

    Dim menu4 As String
    menu4 = Cells(currentRow, 7)

    Dim menu As String
    If menu1 = "N/A" Or menu1 = "" Then
        menu = ""
    Else
        If menu4 = "N/A" Or menu4 = "" Then
            menu = menu1 & " >> " & menu2 & " >> " & menu3
        Else
            menu = menu1 & " >> " & menu2 & " >> " & menu3 & " >> " & menu4
        End If
    End If

    Dim highLight As String
    If menu <> "" Then
        highLight = highLight & "> [!Getting information]" & Chr(13) & Chr(10)
        highLight = highLight & "> Ribbon path please refer to ==**" & menu & "**==" & Chr(13) & Chr(10)
    End If

    Dim i As Long
    Dim proc As String
    Dim VBComp As VBComponent
    Dim objProject As VBIDE.VBProject
    Dim objCode As VBIDE.CodeModule

    Dim codeOfLine As String
    Dim codeOfLineMd As String

    Dim resultStrBas As String
    Dim resultStrMd As String

    Set objProject = ThisWorkbook.VBProject
    For Each VBComp In objProject.VBComponents
        If module = VBComp.Name Then
            ' Find the code module for the project.
            Set objCode = VBComp.CodeModule
            For i = 1 To objCode.CountOfLines
                codeOfLine = objCode.Lines(i, 1)
                ' If Trim(codeOfLine) <> "" Then
                proc = objCode.ProcOfLine(i, vbext_pk_Proc)
                If subb = proc Then
                    resultStrBas = resultStrBas & codeOfLine & Chr(13) & Chr(10)
                    codeOfLine = RTrim(codeOfLine)
                    If Len(codeOfLine) > 0 Then
                        codeOfLineMd = LPad(LTrim(codeOfLine), Len(codeOfLine), "&nbsp;")
                        codeOfLine = LTrim(codeOfLine)
                        codeOfLineMd = Replace(codeOfLineMd, codeOfLine, "`" & codeOfLine & "`")
                    Else
                        codeOfLineMd = Replace(Space(4), " ", "&nbsp;")
                    End If
                    resultStrMd = resultStrMd & codeOfLineMd & Chr(13) & Chr(10)
                End If
                ' End If
            Next i
        End If
    Next VBComp
    Dim beCallArr As Variant
    beCallArr = Split(beCall, Chr(13) & Chr(10))
    Dim j As Integer
    Dim beCallProc As String
    'MsgBox UBound(beCallArr)
    For j = 0 To UBound(beCallArr) - 1
        If EndsWith(CStr(beCallArr(j)), "{S}") Then
            beCallProc = CutStringByStartAndEnd(CStr(beCallArr(j)), "->", "{")
            If InStr(resultStrMd, "Call " & beCallProc) > 0 Then
                resultStrMd = Replace(resultStrMd, "`Call " & beCallProc & "`", "`Call `" & "[`" & beCallProc & "`](" & beCallProc & ")")
            Else
                resultStrMd = Replace(resultStrMd, "`" & beCallProc & " ", "[`" & beCallProc & "`](" & beCallProc & ")` ")
            End If
            GoTo ContinueLoop
        End If
        If EndsWith(CStr(beCallArr(j)), "{F}") Then
            beCallProc = CutStringByStartAndEnd(CStr(beCallArr(j)), "->", "{")
            resultStrMd = Replace(resultStrMd, " " & beCallProc & "(", " `[`" & beCallProc & "`](" & beCallProc & ")`(")
            GoTo ContinueLoop
        End If
ContinueLoop:
    Next j

    If highLight <> "" Then
        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & highLight
    End If

'    If Trim(call1) <> "" Then
'        call1 = Replace(call1, "{S}(", "]]{S}(")
'        call1 = Replace(call1, "{F}(", "]]{F}(")
'        call1 = Replace(call1, "<-", "<-[[")
'        call1 = "# " & "Caller1" & Chr(13) & Chr(10) & call1
'        call1 = Replace(call1, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")
'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call1, Len(call1) - 2)
'    End If
'
'    If Trim(call2) <> "" Then
'        call2 = Replace(call2, "{S}(", "]]{S}(")
'        call2 = Replace(call2, "{F}(", "]]{F}(")
'        call2 = Replace(call2, "<-", "<-[[")
'        call2 = "# " & "Caller2" & Chr(13) & Chr(10) & call2
'        call2 = Replace(call2, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")
'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call2, Len(call2) - 2)
'    End If
'
'    If Trim(call3) <> "" Then
'        call3 = Replace(call3, "{S}(", "]]{S}(")
'        call3 = Replace(call3, "{F}(", "]]{F}(")
'        call3 = Replace(call3, "<-", "<-[[")
'        call3 = "# " & "Caller3" & Chr(13) & Chr(10) & call3
'        call3 = Replace(call3, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")
'        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(call3, Len(call3) - 2)
'    End If

    If Trim(beCall) <> "N/A" Then
        beCall = Replace(beCall, "{S}" & Chr(13) & Chr(10), "]]{S}" & Chr(13) & Chr(10))
        beCall = Replace(beCall, "{F}", "]]{F}")
        beCall = Replace(beCall, "->", "->[[")
        beCall = "# " & "BeCaller" & Chr(13) & Chr(10) & beCall
        beCall = Replace(beCall, Chr(13) & Chr(10), Chr(13) & Chr(10) & "- ")
        resultStrMd = resultStrMd & Chr(13) & Chr(10) & Chr(13) & Chr(10) & Left(beCall, Len(beCall) - 2)
    End If

    Dim folderSrc As String
    Dim folderMd As String
    folderSrc = "C:\SANDBOX\VB_SPACE\GIT_CST_MD\" & Format(Now, "yyyyMMdd") & "\" & module & "\"
    'folderSrc = "C:\SANDBOX\VB_SPACE\GIT_CST_MD\" & Format(Now, "yyyyMMddHHmm") & "\" & module & "\"
    folderMd = "C:\MD_SPACE\" & module & "\"

    CreateFolder folderSrc
    CreateFolder folderMd

    WriteTxt2Tmp resultStrBas, folderSrc & subb & ".bas"
    WriteTxt2Tmp resultStrMd, folderMd & subb & ".md"
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 5
        Cells(currentRow, 47) = "###"
    End If
End Sub

