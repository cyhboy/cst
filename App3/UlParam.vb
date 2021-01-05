
Public Sub UlParam(Hold As Boolean)
    If testing Then Exit Sub
    On Error GoTo ErrorHandler
    
    Call Fold
    
    Dim path As String
    Dim parameter As String
    Dim currentRow As Integer

    path = GetAppDrive() & "\ptty\pscp.exe -i C:\AppFiles\ptty\myPri.ppk"
    currentRow = ActiveCell.Row
    
    Dim hn As String
    hn = Cells(currentRow, 2)
    
    Dim uid As String
    uid = Cells(currentRow, 3)
    Dim pass As String
    pass = Cells(currentRow, 4)
    
    Dim port As String
    port = Cells(currentRow, 7)
    If (port = "") Then
        port = "22"
    End If
    
    If (Len(port) > 5) Then
        port = "22"
    End If
    
    If Not IsNumeric(port) Then
        port = "22"
    End If
    
    Dim fileOrFolder As String
    fileOrFolder = Cells(currentRow, 5)
    
    Dim Length As Integer
    Length = Len(fileOrFolder)
    
    Dim index As Integer
    index = InStrRev(fileOrFolder, "/")
    
    fileOrFolder = Left(fileOrFolder, index)
    
    Dim specialFile As String
    specialFile = Cells(currentRow, 11)
    
    'fileOrFolder = fileOrFolder & specialFile
    
    Dim localFolder As String
    localFolder = Cells(currentRow, 9)
    
    If Not EndsWith(localFolder, "\") Then
        localFolder = localFolder & "\"
    End If
    
    If pass = "" Then
        'uid = Environ$("username")
        
        Dim propsMap As Variant
        Set propsMap = ReadPropertyInAppFiles("identity.ini")

        pass = propsMap("AD_PASSWORD")
    End If
    If pass = "" Then
        parameter = " -sftp -l " & uid & " -P " & port & " -p """ & localFolder & specialFile & """" & " """ & hn & ":" & fileOrFolder & """"
    Else
        parameter = " -sftp -l " & uid & " -pw " & pass & " -P " & port & " -p """ & localFolder & specialFile & """" & " """ & hn & ":" & fileOrFolder & """"
    End If
    'MsgBox path & parameter
    'Exit Sub
    
    ShellRunMax path & parameter
    
ErrorHandler:
    If Err.Number <> 0 Then
        MyMsgBox Err.Number & " " & Err.Description, 30
    End If
    
    If Hold Then
        Dim exeName As String: exeName = ExtractEXE(path)
        While True = IsExeRunning(exeName)
            Sleep 10000
            ShellRun path & "--close"
        Wend
    End If
    
End Sub

