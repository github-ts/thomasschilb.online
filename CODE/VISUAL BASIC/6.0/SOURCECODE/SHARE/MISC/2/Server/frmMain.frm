VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote File Server"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7365
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
   ScaleHeight     =   2490
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin FileServer.FileSender Sender 
      Left            =   5160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      FileTitle       =   "(File title)"
      MainPort        =   4444
      TransferPort    =   5555
   End
   Begin FileServer.FileReceiver Receiver 
      Left            =   4440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      ReceiveDirectory=   ""
   End
   Begin VB.FileListBox Files 
      Height          =   1260
      Left            =   2760
      ReadOnly        =   0   'False
      System          =   -1  'True
      TabIndex        =   16
      Top             =   3360
      Width           =   1815
   End
   Begin VB.DirListBox Dirs 
      Height          =   990
      Left            =   600
      TabIndex        =   15
      Top             =   3720
      Width           =   2055
   End
   Begin VB.DriveListBox Drives 
      Height          =   315
      Left            =   600
      TabIndex        =   14
      Top             =   3360
      Width           =   2055
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   3240
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   7080
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   2280
      MaxLength       =   5
      TabIndex        =   11
      Text            =   "7080"
      Top             =   1800
      Width           =   1455
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   2250
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   423
      Style           =   1
      SimpleText      =   "Status: Server closed."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin FileServer.Button cmdReset 
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Caption         =   "Reset"
      ForeColor       =   -2147483630
   End
   Begin FileServer.Button cmdStart 
      Height          =   285
      Left            =   4920
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      Caption         =   "Start"
      ForeColor       =   -2147483630
   End
   Begin FileServer.Button cmdClose 
      Height          =   285
      Left            =   6120
      TabIndex        =   13
      Top             =   1800
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      Caption         =   "Close"
      ForeColor       =   -2147483630
      Enabled         =   0   'False
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Accept connections on port:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   2025
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00BD662D&
      X1              =   120
      X2              =   7200
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblCurDir 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Current directory:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   1305
   End
   Begin VB.Label lblKBR 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   840
      Width           =   105
   End
   Begin VB.Label lblKBS 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   480
      Width           =   105
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "KB received:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "KB sent:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   600
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Currently connected:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
   Begin VB.Menu mnuSession 
      Caption         =   "&Session"
      Begin VB.Menu mnuOpenServer 
         Caption         =   "Open Server"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuCloseServer 
         Caption         =   "Close Server"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuSep1 
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

Private Sub cmdClose_Click()
On Error Resume Next
sckMain.Close
Receiver.CloseSocket
Sender.CloseSocket
cmdReset_Click
bServerOpen = False
lblCon.Caption = Empty
lblCurDir.Caption = Empty
sCurDir = Empty
sCurCon = Empty
cmdClose.Enabled = False
cmdStart.Enabled = True
StatusBar.SimpleText = "Status: Server closed."
End Sub

Private Sub cmdReset_Click()
lblKBS.Caption = "0"
lblKBR.Caption = "0"
dBytesSent = 0
dBytesRec = 0
End Sub

Private Sub cmdStart_Click()
If Len(txtPort.Text) = 0 Then
    MsgBox "Enter a port to accept connections on", vbCritical, "Local Port Required"
    txtPort.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtPort.Text) Then
    MsgBox "Enter a numeric value for the port", vbCritical, "Invalid Port Value"
    txtPort.SetFocus
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
    Exit Sub
End If
sckMain.LocalPort = txtPort.Text
sckMain.Listen
bServerOpen = True
cmdReset_Click
lblCon.Caption = Empty
lblCurDir.Caption = Empty
sCurDir = Empty
sCurCon = Empty
StatusBar.SimpleText = "Status: Server open."
cmdStart.Enabled = False
cmdClose.Enabled = True
Receiver.Listen
End Sub

Private Sub Form_Load()
cmdStart_Click
End Sub

Private Sub mnuCloseServer_Click()
If cmdClose.Enabled Then cmdClose_Click
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuOpenServer_Click()
If cmdStart.Enabled Then cmdStart_Click
End Sub

Private Sub mnuSession_Click()
mnuOpenServer.Enabled = Not bServerOpen
mnuCloseServer.Enabled = bServerOpen
End Sub

Private Sub Receiver_DataReceived(DataLen As Long)
dBytesRec = dBytesRec + DataLen
lblKBR.Caption = BytesToKB(dBytesRec)
End Sub

Private Sub Receiver_DataSent(DataLen As Long)
dBytesSent = dBytesSent + DataLen
lblKBS.Caption = BytesToKB(dBytesSent)
End Sub

Private Sub Receiver_ReceivedTransferRequest(ByVal FileTitle As String, ByVal FileSize As Double)
If Len(sCurDir) = 0 Then Exit Sub
Receiver.ReceiveDirectory = sCurDir
End Sub

Private Sub Receiver_SocketClosed()
On Error Resume Next
Receiver.CloseSocket
Receiver.Listen
End Sub

Private Sub Receiver_TransferComplete()
Receiver.Reset
End Sub

Private Sub sckMain_Close()
StatusBar.SimpleText = "Status: " & sckMain.RemoteHostIP & " disconnected."
sckMain.Close
sckMain.Listen
sCurCon = Empty
sCurDir = Empty
lblCon.Caption = Empty
lblCurDir.Caption = Empty
End Sub

Private Sub sckMain_ConnectionRequest(ByVal requestID As Long)
sckMain.Close
sckMain.Accept requestID
sCurCon = sckMain.RemoteHostIP
lblCon.Caption = sCurCon
End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
Dim sData As String, sBuff() As String, iLoop As Integer, sTmpCMD As String
sckMain.GetData sData, vbString, bytesTotal
dBytesRec = dBytesRec + bytesTotal
lblKBR.Caption = BytesToKB(dBytesRec)
sBuff() = Split(sData, EOP)
For iLoop = 0 To UBound(sBuff)
    sTmpCMD = sBuff(iLoop)
    If Len(sTmpCMD) > 0 Then
        If Left(sTmpCMD, 3) = "GET" Then
            Call ParseGetDrives(sTmpCMD)
        ElseIf Left(sTmpCMD, 3) = "CHG" Then
            Call ParseChangeDirectory(sTmpCMD)
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

Private Sub sckMain_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
dBytesSent = dBytesSent + bytesSent
lblKBS.Caption = BytesToKB(dBytesSent)
End Sub

Private Sub Sender_TransferComplete()
Sender.Reset
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
' Number only
If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
