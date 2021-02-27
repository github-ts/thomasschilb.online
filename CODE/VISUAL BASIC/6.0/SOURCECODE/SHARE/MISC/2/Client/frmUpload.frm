VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmUpload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Upload File"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
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
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkClose 
      Caption         =   "Close window when upload completes."
      Height          =   255
      Left            =   1560
      TabIndex        =   14
      Top             =   3720
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin RemoteFileManager.FileSender Sender 
      Left            =   2520
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      FileTitle       =   ""
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin RemoteFileManager.Button cmdUpload 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Upload"
      ForeColor       =   -2147483630
   End
   Begin ComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4050
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Status: Idle."
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   " File information "
      Height          =   855
      Left            =   1320
      TabIndex        =   4
      Top             =   1560
      Width           =   5295
      Begin VB.Label lblFS 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label lblFN 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "File size (KB):"
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "File name:"
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
   Begin RemoteFileManager.Button cmdBrowse 
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      ToolTipText     =   " Browse "
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      Caption         =   "..."
      ForeColor       =   -2147483630
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   4215
   End
   Begin RemoteFileManager.Button cmdCancel 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Cancel"
      ForeColor       =   -2147483630
   End
   Begin VB.Label lblKBPS 
      AutoSize        =   -1  'True
      Caption         =   "0 KB/Sec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "File to upload:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Do not close this window until the file transfer is complete."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BD662D&
      Height          =   240
      Left            =   1140
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmUpload.frx":0A02
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
Set objFSO = New FileSystemObject
With CD
    .DialogTitle = "Select a File to Upload"
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
    If Len(.Filename) = 0 Then Exit Sub
    If Not objFSO.FileExists(.Filename) Then
        MsgBox "The file you have selected does not exist", vbCritical, "Invalid File Specified"
        Exit Sub
    ElseIf FileLen(.Filename) = 0 Then
        MsgBox "The file you have selected is empty", vbCritical, "Invalid File Specified"
        Exit Sub
    End If
    txtPath.Text = .Filename
    lblFN.Caption = .FileTitle
    Dim dFileLen As Double
    dFileLen = FileLen(.Filename)
    lblFS.Caption = BytesToKB(dFileLen)
    Sender.FilePath = .Filename
    Sender.FileTitle = .FileTitle
    Bar.Max = dFileLen
End With
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Sender.StopTransfer
Sender.Reset
Unload Me
End Sub

Private Sub cmdUpload_Click()
If Len(txtPath.Text) = 0 Then
    MsgBox "Select a file to upload", vbCritical, "File Required"
    cmdBrowse.SetFocus
    Exit Sub
End If
Bar.Value = Bar.Min
lblKBPS.Caption = "0 KB/Sec"
Sender.RemoteHost = frmMain.sckMain.RemoteHostIP
Sender.Connect
StatusBar.SimpleText = "Status: Connecting..."
End Sub

Private Sub Sender_Connected()
Bar.Value = Bar.Min
StatusBar.SimpleText = "Status: Connected! Negotiating transfer..."
End Sub

Private Sub Sender_KBPS(ByVal KBPS As Long)
lblKBPS.Caption = BytesToKB(KBPS) & " KB/Sec"
End Sub

Private Sub Sender_ProgressUpdate(ByVal BytesSent As Double, ByVal BytesTotal As Double)
On Error Resume Next
Bar.Value = BytesSent
End Sub

Private Sub Sender_SocketError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Sender.CloseSocket
Sender.Reset
StatusBar.SimpleText = "Status: " & ErrorDescription
lblKBPS.Caption = "0 KB/Sec"
End Sub

Private Sub Sender_TransferComplete()
StatusBar.SimpleText = "Status: Upload complete."
Bar.Value = Bar.Max
lblKBPS.Caption = "0 KB/Sec"
Call ChangeDirectory(sCurDir)
Sender.Reset
If chkClose.Value Then Unload Me
End Sub

Private Sub Sender_TransferStarted()
StatusBar.SimpleText = "Status: Uploading file..."
End Sub
