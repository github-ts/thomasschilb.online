VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote File Manager"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   6000
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   5280
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   7080
   End
   Begin RemoteFileManager.Button cmdDelete 
      Height          =   375
      Left            =   8520
      TabIndex        =   12
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Delete File"
      ForeColor       =   -2147483630
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   6180
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status: Not connected."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCurDir 
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5040
      Width           =   8535
   End
   Begin MSComctlLib.ImageList ILDir 
      Left            =   360
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3482
            Key             =   "Dir"
            Object.Tag             =   "Dir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILFile 
      Left            =   4080
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3E94
            Key             =   "FN"
            Object.Tag             =   "FN"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":442E
            Key             =   "FS"
            Object.Tag             =   "FS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVFile 
      Height          =   3495
      Left            =   4320
      TabIndex        =   8
      Top             =   1440
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6165
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ILFile"
      SmallIcons      =   "ILFile"
      ColHdrIcons     =   "ILFile"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File name"
         Object.Width           =   6668
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File size (KB)"
         Object.Width           =   3352
         ImageIndex      =   2
      EndProperty
   End
   Begin MSComctlLib.TreeView TVDir 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ILDir"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.ComboBox cmbDrives 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9975
      Begin RemoteFileManager.Button cmdConnect 
         Height          =   285
         Left            =   8400
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         Caption         =   "Connect"
         ForeColor       =   -2147483630
      End
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   6720
         MaxLength       =   5
         TabIndex        =   4
         Text            =   "7080"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtHost 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Remote port:"
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   5640
         TabIndex        =   3
         Top             =   240
         Width           =   960
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmMain.frx":5DC0
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remote host:"
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin RemoteFileManager.Button cmdRemoveDir 
      Height          =   375
      Left            =   6840
      TabIndex        =   13
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Remove Directory"
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdFileInfo 
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "File Info..."
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdExecute 
      Height          =   375
      Left            =   3480
      TabIndex        =   15
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Execute File"
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdUpload 
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Upload..."
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdDownload 
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Download..."
      ForeColor       =   -2147483630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "<- Select a drive from the list."
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   2400
      TabIndex        =   18
      Top             =   960
      Width           =   2160
   End
   Begin VB.Image imgBack 
      Height          =   240
      Left            =   1920
      MouseIcon       =   "frmMain.frx":67C2
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":6ACC
      ToolTipText     =   " Previous Directory "
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Current directory:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   1305
   End
   Begin VB.Menu mnuSession 
      Caption         =   "&Session"
      Begin VB.Menu mnuConnect 
         Caption         =   "Connect"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuDisconnect 
         Caption         =   "Disconnect"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddressBook 
         Caption         =   "Address Book..."
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuAlwaysOnTop 
         Caption         =   "Always on Top"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuGetHelp 
         Caption         =   "Get Help!"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbDrives_Click()
If Not bCon Then Exit Sub
If Not bChangeDrive Then Exit Sub
If Len(Trim$(cmbDrives.Text)) = 0 Then Exit Sub
sCurDir = cmbDrives.Text & "\"
txtCurDir.Text = sCurDir
TVDir.Nodes.Clear
LVFile.ListItems.Clear
Call ChangeDirectory(cmbDrives.Text)
StatusBar.SimpleText = "Status: Navigating..."
End Sub

Sub cmdConnect_Click()
If cmdConnect.Caption = "Connect" Then
    If Len(txtHost.Text) = 0 Then
        MsgBox "Enter a host to connect to", vbCritical, "Remote Host Required"
        txtHost.SetFocus
        Exit Sub
    ElseIf Len(txtPort.Text) = 0 Then
        MsgBox "Enter a port to connect on", vbCritical, "Remote Port Required"
        txtPort.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(txtPort.Text) Then
        MsgBox "Enter a numeric value for the port", vbCritical, "Invalid Port Value"
        txtPort.SetFocus
        txtPort.SelStart = 0
        txtPort.SelLength = Len(txtPort.Text)
        Exit Sub
    End If
    Call SaveCon
    sckMain.Close
    bCon = False
    Call GoDiscon
    sckMain.Connect txtHost.Text, txtPort.Text
    StatusBar.SimpleText = "Status: Connecting..."
ElseIf cmdConnect.Caption = "Disconnect" Then
    sckMain.Close
    bCon = False
    Call GoDiscon
    StatusBar.SimpleText = "Status: Disconnected."
    cmdConnect.Caption = "Connect"
End If
End Sub

Private Sub cmdDelete_Click()
If Not bCon Then Exit Sub
If Len(sCurDir) = 0 Then Exit Sub
Dim sSel As String
sSel = GetSelectedFile
If Len(sSel) = 0 Then
    MsgBox "Select a file from the list", vbCritical, "File Required"
    Exit Sub
Else
    Dim sFullPath As String
    If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
    sFullPath = sCurDir & sSel
    Call SendDeleteFile(sFullPath)
    StatusBar.SimpleText = "Status: Deleting " & Chr$(34) & sSel & Chr$(34) & "..."
End If
End Sub

Private Sub cmdDownload_Click()
If Not bCon Then Exit Sub
Set objFSO = New FileSystemObject
Dim sSel As String, sExt As String
sSel = GetSelectedFile
If Len(sSel) = 0 Then
    MsgBox "Select a file from the list to download", vbCritical, "File Required"
    Exit Sub
Else
    With CD
        .DialogTitle = "Save File As"
        sExt = objFSO.GetExtensionName(sSel)
        If Len(sExt) = 0 Then Exit Sub
        .Filter = "(*." & sExt & " Files)|*." & sExt
        .Filename = objFSO.GetFileName(sSel)
        .ShowSave
        If Len(.Filename) = 0 Then Exit Sub
        If Len(sCurDir) = 0 Then Exit Sub
        If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
        frmDownload.Receiver.CloseSocket
        frmDownload.ResetTransfer
        frmDownload.Receiver.ReceiveDirectory = objFSO.GetParentFolderName(.Filename)
        frmDownload.Receiver.Listen
        frmDownload.StatusBar.SimpleText = "Status: Negotiating..."
        frmDownload.txtPath.Text = .Filename
        frmDownload.lblFN.Caption = sSel
        frmDownload.Show
        Call SendDownloadFile(sCurDir & sSel)
    End With
End If
End Sub

Function GetSelectedFile() As String
On Error Resume Next
Dim lSel As Long, sSel As String
lSel = LVFile.SelectedItem.Index
If lSel = 0 Then Exit Function
sSel = Trim$(LVFile.ListItems(lSel).Text)
If Len(sSel) = 0 Then Exit Function
GetSelectedFile = sSel
End Function

Private Sub cmdExecute_Click()
If Not bCon Then Exit Sub
If Len(sCurDir) = 0 Then Exit Sub
Dim sSel As String, sExt As String
sSel = GetSelectedFile
If Len(sSel) = 0 Then
    MsgBox "Select a file from the list to execute", vbCritical, "File Required"
    Exit Sub
Else
    Dim sFullPath As String
    If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
    sFullPath = sCurDir & sSel
    Call SendExeFile(sFullPath)
    StatusBar.SimpleText = "Status: Executing " & Chr$(34) & sSel & Chr$(34) & "..."
End If
End Sub

Private Sub cmdFileInfo_Click()
If Not bCon Then Exit Sub
If Len(sCurDir) = 0 Then Exit Sub
Dim sSel As String
sSel = GetSelectedFile
If Len(sSel) = 0 Then
    MsgBox "Select a file from the list", vbCritical, "File Required"
    Exit Sub
Else
    Dim sFullPath As String
    If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
    sFullPath = sCurDir & sSel
    StatusBar.SimpleText = "Status: Getting file information..."
    Call SendGetFileInfo(sFullPath)
End If
End Sub

Private Sub cmdRemoveDir_Click()
If Not bCon Then Exit Sub
If Len(sCurDir) = 0 Then Exit Sub
Dim sSel As String
sSel = GetSelectedDir
If Len(sSel) = 0 Then
    MsgBox "Select a folder from the list", vbCritical, "File Required"
    Exit Sub
Else
    Dim sFullPath As String
    If Not Right(sCurDir, 1) = "\" Then sCurDir = sCurDir & "\"
    sFullPath = sCurDir & sSel
    Call SendRemoveDirectory(sFullPath)
    StatusBar.SimpleText = "Status: Removing directory..."
End If
End Sub

Private Sub cmdUpload_Click()
If Not bCon Then Exit Sub
frmUpload.Show vbModal
End Sub

Private Sub Form_Load()
Call LoadCon
End Sub

Private Sub imgBack_Click()
If Not bCon Then Exit Sub
If Len(sCurDir) = 0 Then Exit Sub
Set objFSO = New FileSystemObject
Dim sTmpDir As String
sTmpDir = objFSO.GetParentFolderName(sCurDir)
If Len(sTmpDir) = 0 Then Exit Sub
If Not Right(sTmpDir, 1) = "\" Then sTmpDir = sTmpDir & "\"
sCurDir = sTmpDir
txtCurDir.Text = sCurDir
Call ChangeDirectory(sCurDir)
StatusBar.SimpleText = "Status: Navigating..."
End Sub

Private Sub mnuAddressBook_Click()
frmAddressBook.Show
End Sub

Private Sub mnuSession_Click()
mnuConnect.Enabled = Not bCon
mnuDisconnect.Enabled = bCon
End Sub

Private Sub sckMain_Close()
bCon = False
Call GoDiscon
StatusBar.SimpleText = "Status: Connection closed/lost."
End Sub

Private Sub sckMain_Connect()
bCon = True
StatusBar.SimpleText = "Status: Connected! Gathering information..."
cmdConnect.Caption = "Disconnect"
Call SendGetDrives
End Sub

Private Sub sckMain_DataArrival(ByVal BytesTotal As Long)
Dim sData As String, sBuff() As String, iLoop As Integer, sTmpCMD As String
sckMain.GetData sData, vbString, BytesTotal
sBuff() = Split(sData, EOP)
For iLoop = 0 To UBound(sBuff)
    sTmpCMD = sBuff(iLoop)
    If Len(sTmpCMD) > 0 Then
        If Left(sTmpCMD, 3) = "GET" Then
            Call ParseGetDrives(sTmpCMD)
            StatusBar.SimpleText = "Status: Information gathered."
        ElseIf Left(sTmpCMD, 3) = "CHG" Then
            Call ParseChangeDirectory(sTmpCMD)
            StatusBar.SimpleText = "Status: Information gathered."
        ElseIf Left(sTmpCMD, 3) = "DOW" Then
            Call ParseDownloadFile(sTmpCMD)
        ElseIf Left(sTmpCMD, 3) = "EXE" Then
            Call ParseExeFile(sTmpCMD)
        ElseIf Left(sTmpCMD, 3) = "FIN" Then
            Call ParseGetFileInfo(sTmpCMD)
        ElseIf Left(sTmpCMD, 3) = "DEL" Then
            Call ParseDeleteFile(sTmpCMD)
        ElseIf Left(sTmpCMD, 3) = "RMD" Then
            Call ParseRemoveDirectory(sTmpCMD)
        End If
    End If
Next iLoop
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
bCon = False
Call GoDiscon
StatusBar.SimpleText = "Status: " & Description
End Sub

Function GetSelectedDir() As String
On Error Resume Next
If Not bCon Then Exit Function
If TVDir.Nodes.Count = 0 Then Exit Function
Dim lSel As Long, sSel As String
lSel = TVDir.SelectedItem.Index
sSel = Trim$(TVDir.Nodes(lSel).Text)
GetSelectedDir = sSel
End Function

Private Sub TVDir_DblClick()
If Not bCon Then Exit Sub
If TVDir.Nodes.Count = 0 Then Exit Sub
Dim sSel As String
sSel = GetSelectedDir
If Len(sSel) = 0 Then
    MsgBox "Select a folder from the list", vbCritical, "Folder Required"
    Exit Sub
Else
    Call ChangeLocalDirectory(sSel)
    StatusBar.SimpleText = "Status: Navigating..."
    Call ChangeDirectory(sCurDir)
End If
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
' Number only
If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 0 Then KeyAscii = 0
End Sub

Private Sub SaveCon()
On Error Resume Next
SaveSetting "RFM", "Main", "Host", txtHost.Text
SaveSetting "RFM", "Main", "Port", txtPort.Text
End Sub

Private Sub DeleteCon()
On Error Resume Next
DeleteSetting "RFM", "Main", "Host"
DeleteSetting "RFM", "Main", "Port"
End Sub

Private Sub LoadCon()
On Error Resume Next
txtHost.Text = GetSetting("RFM", "Main", "Host", Empty)
txtPort.Text = GetSetting("RFM", "Main", "Port", 7080)
End Sub
