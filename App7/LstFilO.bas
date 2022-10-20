
Public Sub LstFilO()
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
                RobotRunByParam "LstFilO"
            End If
        Next curCell
        Exit Sub
    End If

    Dim currentRow As Integer
    currentRow = ActiveCell.Row

    Dim localFolder As String
    Dim keyword As String
    localFolder = Cells(currentRow, 9)
    keyword = Cells(currentRow, 13)

    Dim act As String
    act = Cells(currentRow, 16)

    '    If EndsWith(localFolder, "\") Then
    '        localFolder = Left(localFolder, Len(localFolder) - 1)
    '    End If

    'MsgBox localFolder & keyword

    Dim currFilename As String
    currFilename = Cells(currentRow, 11)

    Dim lstFilename1 As String

    Dim fileList As Variant
    fileList = GetFileList(localFolder & keyword)

    Dim date1 As Date

    Dim dateOld As Date
    dateOld = DateAdd("yyyy", -20, Now)
    'date1 = Cells(currentRow, 12)

    Dim found As Boolean
    found = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim myFileObj As Object
    Dim myFile As Variant
    Dim i As Integer

    For Each myFile In fileList

        If found Then
            Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
            date1 = myFileObj.DateLastModified
            lstFilename1 = myFileObj.Name
            Exit For
        End If

        If i = 0 Then
            Set myFileObj = fso.GetFile(localFolder & CStr(myFile))
            date1 = myFileObj.DateLastModified
            lstFilename1 = myFileObj.Name
            If myFileObj.Name = currFilename Then
                found = True
            End If
        Else
            Set myFileObj = fso.GetFile(localFolder & CStr(myFile))

            If myFileObj.DateLastModified > dateOld Then
                If myFileObj.Name = currFilename Then
                    'If i < 3 Then
                    found = True
                    'End If
                End If
            End If

        End If

        i = i + 1
    Next myFile

    Set fso = Nothing

    Cells(currentRow, 11) = lstFilename1

    If InStr(act, "Em") = 0 Then
        Cells(currentRow, 12) = date1
    End If
ErrorHandler:
    If Err.Number <> 0 Then
        'MyMsgBox Err.Number & " " & Err.Description, 30
    End If
End Sub

