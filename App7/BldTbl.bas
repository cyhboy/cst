
Public Sub BldTbl()
    If testing Then
        Exit Sub
    End If
    'On Error GoTo ErrorHandler

    Const ForReading = 1, ForWriting = 2
    Dim dbname As String
    Dim tblname As String

    If InStr(ActiveSheet.Name, ".") = 0 Then
        dbname = "common_data"
        tblname = ActiveSheet.Name
    Else
        dbname = Split(ActiveSheet.Name, ".")(0)
        tblname = Split(ActiveSheet.Name, ".")(UBound(Split(ActiveSheet.Name, ".")))
    End If

    Dim sql1 As String

    sql1 = "drop table " & tblname

    Dim filePath As String
    filePath = FnLstFil("C:\BAK\" & ActiveSheet.Name & "_*.txt")

    Dim firstLine As String
    firstLine = FnGetFileLine(filePath, 1)
    firstLine = Left(firstLine, Len(firstLine) - 1)
    firstLine = Right(firstLine, Len(firstLine) - 1)

    Dim firstLineAry As Variant
    firstLineAry = Split(firstLine, "','")

    Dim sql2 As String
    sql2 = "CREATE TABLE " & tblname & " ("

    Dim i As Integer
    Dim k As Integer

    For i = 0 To UBound(firstLineAry)
        sql2 = sql2 & firstLineAry(i) & " LONGTEXT,"
    Next i

    'sql2 = sql2 & "remark LONGTEXT"
    sql2 = Left(sql2, Len(sql2) - 1)
    sql2 = sql2 & ")"

    MsgBox sql2


    Dim strPath As String
    strPath = "M:\AppFiles\" & dbname & ".accdb"
    'strPath = "C:\AppFiles\" & dbname & ".accdb"

    If Dir(strPath) = "" Then
        strPath = Replace(strPath, Left(strPath, 2), "C:")
    End If

    Dim db As DAO.Database
    Dim rst As DAO.Recordset

    Set db = OpenDatabase(strPath)

    On Error Resume Next
    'conn.Execute sql1
    db.Execute sql1
    On Error GoTo 0
    'MsgBox sql2
    db.Execute sql2
    'conn.Execute sql2

    Dim dataLineAry As Variant

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, ForReading)

    Dim lineNo As Long
    lineNo = 0
    Dim dataLine As String
    Set rst = db.OpenRecordset(tblname)
    Do Until ts.AtEndOfStream
        lineNo = lineNo + 1
        dataLine = ts.readline
        If lineNo > 1 Then
            dataLine = Left(dataLine, Len(dataLine) - 1)
            dataLine = Right(dataLine, Len(dataLine) - 1)
            dataLineAry = Split(dataLine, "','")


            rst.AddNew
            For k = 0 To UBound(firstLineAry)
                rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
            Next k

            rst.Update
        End If
    Loop

    Set ts = Nothing
    Set fso = Nothing
    rst.Close
    db.Close

    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub

