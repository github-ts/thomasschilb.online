Attribute VB_Name = "ModSession"
Option Explicit

Public Const DELIM As String = "///"
Public Const EOP As String = "**"

Global bCon As Boolean
Global bChangeDrive As Boolean

Global lBytesSent As Long
Global lBytesRec As Long

Global sCurDir As String

Global objFSO As FileSystemObject

Public Sub GoDiscon()
With frmMain
    .TVDir.Nodes.Clear
    .cmbDrives.Clear
    .LVFile.ListItems.Clear
    .txtCurDir.Text = Empty
    .cmdConnect.Caption = "Connect"
    sCurDir = Empty
End With
End Sub

Public Sub SendGetDrives()
' "GET"/"Drives"
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "GET" & DELIM & "Drives"
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseGetDrives(ByVal Data As String)
' "GET"/"Drives"/Drive1|Drive2|Drive3
On Error Resume Next
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
Dim sDrives() As String, iLoop As Integer
sDrives() = Split(sBuff(2), "|")
With frmMain.cmbDrives
    .Clear
    For iLoop = 0 To UBound(sDrives)
        If Len(sDrives(iLoop)) > 0 Then
            .AddItem sDrives(iLoop)
        End If
    Next iLoop
    bChangeDrive = False
    .Text = .List(0)
    bChangeDrive = True
End With
End Sub

Public Sub ChangeDirectory(ByVal NewDirectory As String)
' "CHG"/Directory
If Not bCon Then Exit Sub
If Len(NewDirectory) = 0 Then Exit Sub
Dim sPacket As String
sPacket = "CHG" & DELIM & NewDirectory
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseChangeDirectory(ByVal Data As String)
On Error Resume Next
' Successfull: "CHG"/"Directory"/D:Folder1|F:File1*123|F:File2*123
' Folder doesn't exist: "CHG"/"Directory"/"NonExist"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(2) = "NonExist" Then
    frmMain.StatusBar.SimpleText = "Status: Error changing directory; directory does not exist."
    Exit Sub
Else
    sCurDir = sBuff(1)
    frmMain.txtCurDir.Text = sCurDir
    Dim sDir() As String, sType() As String, lLoop As Long, sTmpDir As String, sSize() As String
    sDir() = Split(sBuff(2), "|")
    With frmMain
        .TVDir.Nodes.Clear
        .LVFile.ListItems.Clear
        For lLoop = 0 To UBound(sDir)
            sTmpDir = sDir(lLoop)
            If Len(sTmpDir) > 0 Then
                sType() = Split(sTmpDir, "*?*")
                If sType(0) = "D" Then
                    .TVDir.Nodes.Add , , sType(1), sType(1), "Dir"
                ElseIf sType(0) = "F" Then
                    sSize() = Split(sType(1), "*")
                    .LVFile.ListItems.Add , , sSize(0), , "FN"
                    .LVFile.ListItems(.LVFile.ListItems.Count).ListSubItems.Add , , BytesToKB(CDbl(sSize(1)))
                    .LVFile.ListItems(.LVFile.ListItems.Count).ListSubItems(1).Bold = True
                End If
            End If
        Next lLoop
        .LVFile.Refresh
    End With
End If
End Sub

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
On Error Resume Next
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

Public Function ChangeLocalDirectory(ByVal NextFolder As String)
If Len(sCurDir) = 0 Then Exit Function
If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
sCurDir = sCurDir & NextFolder & "\"
frmMain.txtCurDir.Text = sCurDir
End Function

Public Sub SendDownloadFile(ByVal FilePath As String)
' "DOW"/File path
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "DOW" & DELIM & FilePath
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseDownloadFile(ByVal Data As String)
' File is empty: "DOW"/"Empty"
' File doesn't exist: "DOW"/"NonExist"
' Ready: "DOW"/"Ready"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
With frmMain.StatusBar
    If sBuff(1) = "Empty" Then
        .SimpleText = "Status: Unable to download file; file is empty."
        frmDownload.StatusBar.SimpleText = .SimpleText
    ElseIf sBuff(1) = "NonExist" Then
        .SimpleText = "Status: Unable to download file; file doesn't exist."
        frmDownload.StatusBar.SimpleText = .SimpleText
    End If
End With
End Sub

Public Sub SendExeFile(ByVal FilePath As String)
' "EXE"/File path
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "EXE" & DELIM & FilePath
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseExeFile(ByVal Data As String)
    'Error: "EXE"/"Error"/Error description
    'Doesn't exist: "EXE"/"NonExist"
    'Executed: "EXE"/"Executed"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
With frmMain
    If sBuff(1) = "NonExist" Then
        .StatusBar.SimpleText = "Status: Unable to execute file; file does not exist."
    ElseIf sBuff(1) = "Executed" Then
        .StatusBar.SimpleText = "Status: File executed."
    ElseIf sBuff(1) = "Error" Then
        .StatusBar.SimpleText = "Status: Error executing file - " & sBuff(2)
    End If
End With
End Sub

Public Sub SendGetFileInfo(ByVal FilePath As String)
' "FIN"/File path
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "FIN" & DELIM & FilePath
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseGetFileInfo(ByVal Data As String)
' "FIN"/File name/File path/File size
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "NonExist" Then
    frmMain.StatusBar.SimpleText = "Status: Error getting file information; file does not exist."
    Exit Sub
ElseIf sBuff(1) = "Empty" Then
    frmMain.StatusBar.SimpleText = "Status: Error getting file information; file is empty."
    Exit Sub
End If
Dim dTmpBytes As Double
With frmFileInfo
    .Caption = "File Information (" & sBuff(1) & ")"
    .lblFN.Caption = sBuff(1)
    .lblFSB.Caption = sBuff(3)
    .txtPath.Text = sBuff(2)
    dTmpBytes = Val(sBuff(3))
    .lblFSK.Caption = BytesToKB(dTmpBytes)
    .Show
End With
End Sub

Public Sub SendDeleteFile(ByVal FilePath As String)
' "DEL"/"File"/File path
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "DEL" & DELIM & "File" & DELIM & FilePath
frmMain.sckMain.SendData sPacket & EOP
End Sub

Public Sub ParseDeleteFile(ByVal Data As String)
    ' Error "DEL"/"File"/"Error"/Error description
    ' File doesn't exist: "DEL"/"File"/"NonExist"
    ' Successful: "DEL"/"File"/"Success"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(2) = "NonExist" Then
    frmMain.StatusBar.SimpleText = "Status: Error deleting file; file does not exist."
ElseIf sBuff(2) = "Success" Then
    frmMain.StatusBar.SimpleText = "Status: File deleted."
ElseIf sBuff(2) = "Error" Then
    frmMain.StatusBar.SimpleText = "Status: Error deleting flie; " & sBuff(2)
End If
Call ChangeDirectory(sCurDir)
End Sub

Public Sub SendRemoveDirectory(ByVal Directory As String)
' "RMD"/Directory
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "RMD" & DELIM & Directory
frmMain.sckMain.SendData sPacket
End Sub

Public Sub ParseRemoveDirectory(ByVal Data As String)
    ' Path doesn't exist: "RMD"/"NonExist"
    ' Error: "RMD"/"Error"/Error description
    ' Removed: "RMD"/"Removed"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "NonExist" Then
    frmMain.StatusBar.SimpleText = "Status: Unable to remove directory; directory does not exist."
ElseIf sBuff(1) = "Removed" Then
    frmMain.StatusBar.SimpleText = "Status: Directory removed"
ElseIf sBuff(1) = "Error" Then
    frmMain.StatusBar.SimpleText = "Status: Error removing directory; " & sBuff(2)
End If
Call ChangeDirectory(sCurDir)
End Sub
