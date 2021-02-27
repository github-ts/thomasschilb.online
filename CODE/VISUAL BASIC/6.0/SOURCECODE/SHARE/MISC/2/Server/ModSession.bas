Attribute VB_Name = "ModSession"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1

Public Const EOP As String = "**"
Public Const DELIM  As String = "///"

Global bServerOpen As Boolean

Global sCurCon As String
Global sCurDir As String

Global dBytesSent As Double
Global dBytesRec As Double

Global objFSO As FileSystemObject

Public Sub ParseGetDrives(ByVal Data As String)
' From client: "GET"/"Drives"
' From server: "GET"/"Drives"/Drive1|Drive2|Drive3
Dim sBuff() As String, sPacket As String, sTmp As String, iLoop As Integer
sBuff() = Split(Data, DELIM)
With frmMain.Drives
    .Refresh
    If .ListCount = 0 Then Exit Sub
    For iLoop = 0 To .ListCount - 1
        If Len(.List(iLoop)) > 0 Then
            sTmp = sTmp & .List(iLoop) & "|"
        End If
    Next iLoop
End With
If Len(sTmp) > 0 And Right(sTmp, 1) = "|" Then sTmp = Mid(sTmp, 1, Len(sTmp) - 1)
sPacket = "GET" & DELIM & "Drives" & DELIM & sTmp
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseChangeDirectory(ByVal Data As String)
' From client: "CHG"/Directory
' From server: "CHG"/"Directory"/D:Folder1|F:File1*123|F:File2*123
Dim sBuff() As String, sPacket As String, sTmpDir As String, sTmpFile As String, sDirCont As String
Dim iLoop As Integer, iRet As Integer, sTmpErr As String, lTmpError As Long, lFileLen As Long
sBuff() = Split(Data, DELIM)
'iRet = ChangeDirectoryLists(sBuff(1), lTmpError, sTmpErr)
Set objFSO = New FileSystemObject
If Not objFSO.FolderExists(sBuff(1)) Then
    sPacket = "CHG" & DELIM & "Directory" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
End If
sCurDir = sBuff(1)
If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
frmMain.Dirs.Path = sCurDir
frmMain.Files.Path = sCurDir
frmMain.Dirs.Refresh
frmMain.Files.Refresh
frmMain.lblCurDir.Caption = sCurDir
With frmMain.Dirs
    For iLoop = 0 To .ListCount - 1
        If Len(.List(iLoop)) > 0 Then
            sTmpDir = sTmpDir & "D*?*" & GetCurrentDirectory(.List(iLoop)) & "|"
        End If
    Next iLoop
End With
iLoop = 0
With frmMain.Files
    For iLoop = 0 To .ListCount - 1
        If Len(.List(iLoop)) > 0 Then
            sTmpFile = sTmpFile & "F*?*" & .List(iLoop) & "*" & FileLen(sCurDir & .List(iLoop)) & "|"
        End If
    Next iLoop
End With
If Len(sTmpDir) > 0 And Right(sTmpDir, 1) = "|" Then sTmpDir = Mid(sTmpDir, 1, Len(sTmpDir) - 1)
sDirCont = sTmpDir & sTmpFile
sPacket = "CHG" & DELIM & sCurDir & DELIM & sDirCont
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Function ChangeDirectoryLists(ByVal sNewDirectory As String, ByRef ErrorNumber As Long, ByRef ErrorDescription As String) As Integer
'1 = successfull
'2 = error
On Error GoTo ErrorHandler
Set objFSO = New FileSystemObject
If Not objFSO.FolderExists(sNewDirectory) Then
    ErrorNumber = 1
    ErrorDescription = "Folder doesn't exist."
    ChangeDirectoryLists = 2
    Exit Function
End If
With frmMain
    .Dirs.Path = sNewDirectory
    .Files.Path = sNewDirectory
    .Dirs.Refresh
    .Files.Refresh
End With
Exit Function
ErrorHandler:
    ErrorNumber = Err.Number
    ErrorDescription = Err.Description
    ChangeDirectoryLists = 2
Exit Function
End Function

Public Function GetAfter(ByVal Text As String, ByVal AfterCharacter As String) As String
On Error GoTo ErrorHandler
Dim lStart As Long
lStart = InStr(1, Text, AfterCharacter)
If lStart > 0 Then
    GetAfter = Mid(Text, lStart + 1)
End If
Exit Function
ErrorHandler:
End Function

Public Function BytesToKB(ByVal Bytes As Double) As String
Dim dRet As Double, sRet As String, sAfter As String
dRet = Format(Bytes / 1024, "####################.##")
sAfter = GetAfter(Str(dRet), ".")
If Len(sAfter) = 0 Then
    sRet = Str$(Replace$(dRet, ".", Empty))
Else
    sRet = Str$(dRet)
End If
BytesToKB = sRet
End Function

Public Function GetCurrentDirectory(ByVal DirPath As String) As String
On Error Resume Next
Dim sBuff() As String
sBuff() = Split(DirPath, "\")
GetCurrentDirectory = sBuff(UBound(sBuff))
End Function

Public Sub ParseDownloadFile(ByVal Data As String)
' From client: "DOW"/File path
' From server:
    'File is empty: "DOW"/"Empty"
    'Ready: "DOW"/"Ready"/File name/File size (connect to client)
Dim sBuff() As String, sPacket As String: sBuff() = Split(Data, DELIM)
Set objFSO = New FileSystemObject
If Not objFSO.FileExists(sBuff(1)) Then
    sPacket = "DOW" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
ElseIf FileLen(sBuff(1)) = 0 Then
    sPacket = "DOW" & DELIM & "Empty"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
Else
    sPacket = "DOW" & DELIM & "Ready"
    frmMain.sckMain.SendData sPacket & EOP
    With frmMain
        .Sender.CloseSocket
        .Sender.RemoteHost = .sckMain.RemoteHostIP
        .Sender.FilePath = sBuff(1)
        .Sender.FileTitle = objFSO.GetFileName(sBuff(1))
        .Sender.Connect
    End With
End If
End Sub

Public Sub ParseExeFile(ByVal Data As String)
' From client: "EXE"/File path
' From server:
    'Error: "EXE"/"Error"/Error description
    'Doesn't exist: "EXE"/"NonExist"
    'Executed: "EXE"/"Executed"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
Dim sPacket As String
On Error GoTo ErrorHandler
Set objFSO = New FileSystemObject
If Not objFSO.FileExists(sBuff(1)) Then
    sPacket = "EXE" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
Else
    Dim lRet As Long
    lRet = ShellExecute(frmMain.hwnd, "open", sBuff(1), vbNullString, vbNullString, SW_SHOWNORMAL)
    sPacket = "EXE" & DELIM & "Executed"
    frmMain.sckMain.SendData sPacket & EOP
End If
Exit Sub
ErrorHandler:
    sPacket = "EXE" & DELIM & "Error" & DELIM & Err.Description
    frmMain.sckMain.SendData sPacket & EOP
Exit Sub
End Sub

Public Sub ParseGetFileInfo(ByVal Data As String)
' From client: "FIN"/File path
' From server:
    ' File doesn't exist: "FIN"/"NonExist"
    ' File is empty: "FIN"/"Empty"
    ' File information: "FIN"/File name/File path/File size
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
Dim sPacket As String
Set objFSO = New FileSystemObject
If Not objFSO.FileExists(sBuff(1)) Then
    sPacket = "FIN" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
ElseIf FileLen(sBuff(1)) = 0 Then
    sPacket = "FIN" & DELIM & "Empty"
    frmMain.sckMain.SendData sPacket & EOP
Else
    sPacket = "FIN" & DELIM & objFSO.GetFileName(sBuff(1)) & DELIM & sBuff(1) & DELIM & FileLen(sBuff(1))
    frmMain.sckMain.SendData sPacket & EOP
End If
End Sub

Public Sub ParseDeleteFile(ByVal Data As String)
' From client: "DEL"/"File"/File path
' From server:
    ' Error "DEL"/"File"/"Error"/Error description
    ' File doesn't exist: "DEL"/"File"/"NonExist"
    ' Successful: "DEL"/"File"/"Success"
On Error GoTo ErrorHandler:
Dim sBuff() As String, sPacket As String: sBuff() = Split(Data, DELIM)
Set objFSO = New FileSystemObject
If Not objFSO.FileExists(sBuff(2)) Then
    sPacket = "DEL" & DELIM & "File" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
Else
    Kill sBuff(2)
    sPacket = "DEL" & DELIM & "File" & DELIM & "Success"
    frmMain.sckMain.SendData sPacket & EOP
End If
Exit Sub
ErrorHandler:
    sPacket = "DEL" & DELIM & "File" & DELIM & "Error" & DELIM & Err.Description
    frmMain.sckMain.SendData sPacket & EOP
Exit Sub
End Sub

Public Sub ParseRemoveDirectory(ByVal Data As String)
' From client: "RMD"/Directory
' From server:
    ' Path doesn't exist: "RMD"/"NonExist"
    ' Error: "RMD"/"Error"/Error description
    ' Removed: "RMD"/"Removed"
Dim sBuff() As String, sPacket As String
On Error GoTo ErrorHandler
sBuff() = Split(Data, DELIM)
Set objFSO = New FileSystemObject
If Not objFSO.FolderExists(sBuff(1)) Then
    sPacket = "RMD" & DELIM & "NonExist"
    frmMain.sckMain.SendData sPacket & EOP
    Exit Sub
Else
    RmDir sBuff(1)
    sPacket = "RMD" & DELIM & "Removed"
    frmMain.sckMain.SendData sPacket & EOP
End If
Exit Sub
ErrorHandler:
    sPacket = "RMD" & DELIM & "Error" & DELIM & Err.Description
    frmMain.sckMain.SendData sPacket & EOP
Exit Sub
End Sub
