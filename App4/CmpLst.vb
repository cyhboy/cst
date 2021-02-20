
Public Sub CmpLst()
    If testing Then Exit Sub
    
    Dim fileName1 As String
    Dim fileName2 As String
    
    Dim currentSheetName As String
    currentSheetName = ActiveSheet.Name
    
    Dim date1 As Date
    Dim date2 As Date
    
    date1 = DateAdd("yyyy", -5, Now)
    date2 = DateAdd("yyyy", -6, Now)
    
    Dim fso As Object
    Dim mysource As Object
    Dim myFile As Object
    Dim myFolder As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set mysource = fso.getfolder(GetBakDrive())
    For Each myFile In mysource.Files
        If myFile.Name Like currentSheetName & "_*" Then
            If myFile.DateLastModified > date1 Then
                date2 = date1
                date1 = myFile.DateLastModified
                
                fileName2 = fileName1
                fileName1 = myFile.path
                
                GoTo continue_loop
            ElseIf myFile.DateLastModified > date2 Then
                date2 = myFile.DateLastModified
                fileName2 = myFile.path
                GoTo continue_loop
            End If
        End If
continue_loop:
        'MsgBox date1
        'MsgBox date2
    Next

    If fileName1 = "" And fileName2 = "" Then
        'MsgBox "go to folder"
        date1 = DateAdd("yyyy", -5, Now)
        date2 = DateAdd("yyyy", -6, Now)
        
        For Each myFolder In mysource.SubFolders
            If myFolder.Name Like currentSheetName & "_*" Then
                If myFolder.DateLastModified > date1 Then
                    date2 = date1
                    date1 = myFolder.DateLastModified
                    
                    fileName2 = fileName1
                    fileName1 = myFolder.path
                    
                    GoTo continue_folder_loop
                ElseIf myFolder.DateLastModified > date2 Then
                    date2 = myFolder.DateLastModified
                    fileName2 = myFolder.path
                    GoTo continue_folder_loop
                End If
            End If
continue_folder_loop:
            'MsgBox date1
            'MsgBox date2
        Next
    End If

    Set fso = Nothing
 
    Dim path As String
    Dim parameter As String
    path = """" & GetAppDrive() & "\Beyond Compare 3\BCompare.exe" & """"
    parameter = " " & """" & fileName2 & """" & " " & """" & fileName1 & """"

    ShellRun path & parameter
    
End Sub

