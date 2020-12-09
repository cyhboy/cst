
Public Sub RstCf(Optional control As IRibbonControl)
    If testing Then Exit Sub
    Cells.FormatConditions.Delete
    MyQuestionBox "Clean Conditional Format Only? ", "No", "Yes", 5
    If confirmation = "Yes" Then Exit Sub
    Cells.Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=CELL(""row"")=ROW()"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.599963377788629
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub


