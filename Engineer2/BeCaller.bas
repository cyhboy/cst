
Public Sub BeCaller()

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
                RobotRunByParam "BeCaller"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row
    Dim module As String
    module = Cells(currentRow, 1)
    Dim procc As String
    procc = Cells(currentRow, 2)

    Dim becallerList As String

    Dim i As Long
    Dim proc As String
    Dim VBComp As VBComponent
    Dim objProject As VBIDE.VBProject
    Dim objCode As VBIDE.CodeModule

    Dim codeOfLine As String

    Dim resultStrBas As String
    Dim callee As String
    Dim firstRow As Boolean
    'firstRow = False
    Dim proccRow As Integer
    proccRow = 0

    Dim proccStr As String

    Set objProject = ThisWorkbook.VBProject
    For Each VBComp In objProject.VBComponents
        If module = VBComp.Name Then
            ' Find the code module for the project.
            Set objCode = VBComp.CodeModule
            For i = 1 To objCode.CountOfLines
                codeOfLine = objCode.Lines(i, 1)
                If Trim(codeOfLine) <> "" And False = StartsWith(Trim(codeOfLine), "'") Then
                    proc = objCode.ProcOfLine(i, vbext_pk_Proc)
                    If procc = proc Then
                        proccRow = proccRow + 1
                        If Not firstRow Then
                            If Trim(codeOfLine) = "" Then
                                firstRow = False
                            Else
                                proccStr = SearchRegxKwInStr(codeOfLine, "^P[^ ]+ ([^ ]+)")
                                firstRow = True
                            End If
                        Else
                            'MsgBox Trim(codeOfLine)


                            If MatchRegx(Trim(codeOfLine), "^Call ([^ ]+)$") Then
                                'MsgBox codeOfLine
                                callee = SearchRegxKwInStr(Trim(codeOfLine), "^Call ([^ ]+)$")
                                If becallerList = "" Then
                                    becallerList = procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{S}"
                                Else
                                    becallerList = becallerList & Chr(13) & Chr(10) & procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{S}"
                                End If
                                GoTo ContinueLoop
                            End If

                            If MatchRegx(Trim(codeOfLine), "^([A-Z][A-Z|a-z|0-9]+) [^<>=,]+(, [^<>=,]+)*$") Then
                                'MsgBox codeOfLine
                                callee = SearchRegxKwInStr(Trim(codeOfLine), "^([A-Z][A-Z|a-z|0-9]+) [^<>=,]+(, [^<>=,]+)*$")
                                If callee <> "End" And callee <> "Exit" And callee <> "Next" And callee <> "With" And callee <> "Sleep" And callee <> "MsgBox" And callee <> "Print" And callee <> "Close" And callee <> "Shell" And callee <> "GoTo" And callee <> "While" And callee <> "ReDim" And callee <> "Kill" And callee <> "Case" And callee <> "If" And callee <> "On" And callee <> "Set" And callee <> "Dim" And callee <> "For" And callee <> "ElseIf" And callee <> "Do" And callee <> "Open" And callee <> "Select" And callee <> "Line" Then
                                    If becallerList = "" Then
                                        becallerList = procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{S}"
                                    Else
                                        becallerList = becallerList & Chr(13) & Chr(10) & procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{S}"
                                    End If
                                End If
                                GoTo ContinueLoop
                            End If

                            'If InStr(codeOfLine, " = " & subName & "(") > 0 _
                            Or InStr(codeOfLine, " <> " & subName & "(") > 0 _
                            Or InStr(codeOfLine, " & " & subName & "(") > 0 _
                            Or InStr(codeOfLine, "(" & subName & "(") > 0 _
                            Or InStr(codeOfLine, " - " & subName & "(") > 0 _
                            Or InStr(codeOfLine, ", " & subName & "(") > 0 Then
                            If MatchRegx(Trim(codeOfLine), "[=|>|<|&|-|,|+|f] ([A-Z][^ '\.,#\$\(=\\]+)\([^,]*(, [^,]+)*\)") Then
                                callee = SearchRegxKwInStr(Trim(codeOfLine), "[=|>|<|&|-|,|+|f] ([A-Z][^ '\.,#\$\(=\\]+)\([^,]*(, [^,]+)*\)")
                                If callee <> "Cells" And callee <> "InStrRev" And callee <> "Left" And callee <> "Len" And callee <> "Split" _
                                    And callee <> "CreateObject" And callee <> "Environ" And callee <> "DateAdd" And callee <> "CStr" _
                                    And callee <> "UCase" And callee <> "Trim" And callee <> "Right" And callee <> "Replace" And callee <> "RGB" And callee <> "Now" And callee <> "Chr" _
                                    And callee <> "Mid" And callee <> "Array" And callee <> "FreeFile" And callee <> "Format" And callee <> "InStr" And callee <> "UBound" _
                                    And callee <> "Dir" And callee <> "Sheets" And callee <> "GetObject" And callee <> "OpenDatabase" And callee <> "IsNumeric" _
                                    And callee <> "CDbl" And callee <> "LBound" And callee <> "TimeSerial" And callee <> "Int" And callee <> "CInt" _
                                    And callee <> "Join" And callee <> "RTrim" And callee <> "LTrim" And callee <> "IIf" And callee <> "Range" And callee <> "MsgBox" _
                                    And callee <> "FileLen" And callee <> "Chr" And callee <> "Chr" And callee <> "Chr" And callee <> "Chr" Then
                                    If becallerList = "" Then
                                        becallerList = procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{F}"
                                    Else
                                        becallerList = becallerList & Chr(13) & Chr(10) & procc & "{" & Left(proccStr, 1) & "}" & "(" & proccRow & ")" & "->" & callee & "{F}"
                                    End If
                                End If
                                GoTo ContinueLoop
                            End If

                        End If

                    End If
                End If
ContinueLoop:
            Next i
        End If
    Next VBComp
    If becallerList <> "" Then
        Cells(currentRow, 13) = becallerList & Chr(13) & Chr(10)
    Else
        Cells(currentRow, 13) = "N/A"
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 5
        Cells(currentRow, 47) = "###"
    End If
End Sub

