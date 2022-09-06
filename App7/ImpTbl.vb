
Public Sub ImpTbl()
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

    Dim filePath As String
    filePath = FnLstFil("C:\BAK\" & ActiveSheet.Name & "_*.txt")

    Dim firstLine As String
    firstLine = FnGetFileLine(filePath, 1)
    firstLine = Left(firstLine, Len(firstLine) - 1)
    firstLine = Right(firstLine, Len(firstLine) - 1)

    Dim firstLineAry As Variant
    firstLineAry = Split(firstLine, "','")


    Dim strPath As String
    strPath = "M:\AppFiles\" & dbname & ".accdb"
    'strPath = "C:\AppFiles\" & dbname & ".accdb"

    If Dir(strPath) = "" Then
        strPath = Replace(strPath, Left(strPath, 2), "C:")
    End If

    Dim db As DAO.Database

    Dim rst As DAO.Recordset

    Set db = OpenDatabase(strPath)

    Dim dataLineAry As Variant
    Dim sql4 As String
    Dim sql5 As String
    Dim sql6 As String

    Dim sql7 As String
    'sql7 = "delete from " & tblname

    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(filePath, ForReading)

    Dim lineNo As Long
    lineNo = 0
    Dim dataLine As String
    Dim cond As String
    Dim setVal As String

    Dim i As Long
    Dim j As Long
    Dim k As Long

    If UBound(firstLineAry) < 14 Then
        'db.Execute sql7

        Do Until ts.AtEndOfStream
            lineNo = lineNo + 1
            dataLine = ts.readline
            If lineNo > 1 Then
                dataLine = Left(dataLine, Len(dataLine) - 1)
                dataLine = Right(dataLine, Len(dataLine) - 1)
                dataLineAry = Split(dataLine, "','")


                cond = " where "
                For i = 0 To 1
                    cond = cond & firstLineAry(i) & "='" & dataLineAry(i) & "' and "
                Next i

                cond = cond & "1=1"

                sql4 = "select * from " & tblname & cond
                sql7 = "delete from " & tblname & cond

                Set rst = db.OpenRecordset(sql4)
                If rst.RecordCount = 1 Then
                    rst.Edit
                    For k = 0 To UBound(firstLineAry)
                        rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                    Next k
                    rst.Update
                ElseIf rst.RecordCount = 0 Then
                    rst.AddNew
                    For k = 0 To UBound(firstLineAry)
                        rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                    Next k
                    rst.Update
                Else
                    db.Execute sql7
                    rst.AddNew
                    For k = 0 To UBound(firstLineAry)
                        rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                    Next k
                    rst.Update
                End If
            End If
        Loop
    Else

        Do Until ts.AtEndOfStream
            lineNo = lineNo + 1
            dataLine = ts.readline
            If lineNo > 1 Then

                dataLine = Left(dataLine, Len(dataLine) - 1)
                dataLine = Right(dataLine, Len(dataLine) - 1)
                dataLineAry = Split(dataLine, "','")


                cond = " where "
                For i = 13 To 14
                    cond = cond & firstLineAry(i) & "='" & dataLineAry(i) & "' and "
                Next i

                cond = cond & "1=1"

                sql4 = "select * from " & tblname & cond
                sql6 = "delete from " & tblname & cond

                Set rst = db.OpenRecordset(sql4)

                If rst.RecordCount = 1 Then
                    rst.Edit
                    For k = 0 To UBound(firstLineAry)
                        If k <> 13 And k <> 14 Then
                            rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                        End If
                    Next k
                    rst.Update
                ElseIf rst.RecordCount = 0 Then
                    rst.AddNew
                    For k = 0 To UBound(firstLineAry)
                        rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                    Next k
                    rst.Update
                Else
                    db.Execute sql6
                    rst.AddNew
                    For k = 0 To UBound(firstLineAry)
                        rst(firstLineAry(k)) = Replace(dataLineAry(k), "\n", vbCrLf)
                    Next k
                    rst.Update
                End If

            End If
        Loop
    End If

    Set ts = Nothing
    Set fso = Nothing
    rst.Close
    db.Close

    MyMsgBox "Done Import", 5

    'ErrorHandler:
    '    If Err.Number <> 0 Then
    '        MyMsgBox Err.Number & " " & Err.Description, 30
    '    End If
End Sub

