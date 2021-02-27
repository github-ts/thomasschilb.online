VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmDownload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Download File"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
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
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin RemoteFileManager.FileReceiver Receiver 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      MainPort        =   4444
      TransferPort    =   5555
      ReceiveDirectory=   ""
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   3825
      Width           =   6885
      _ExtentX        =   12144
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
   Begin VB.CheckBox chkClose 
      Caption         =   "Close window when download completes."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      Caption         =   " File information "
      Height          =   855
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   5295
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "File name:"
         ForeColor       =   &H00BD662D&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
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
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
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
         TabIndex        =   4
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   5295
   End
   Begin ComctlLib.ProgressBar Bar 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin RemoteFileManager.Button cmdCancel 
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   3360
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
      Top             =   3000
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Saving file to:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   990
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
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmDownload.frx":0A02
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub ResetTransfer()
Bar.Value = Bar.Min
Bar.Max = 100
lblKBPS.Caption = "0 KB/Sec"
lblFN.Caption = "-"
lblFS.Caption = "-"
txtPath.Text = Empty
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
Receiver.StopTransfer
Receiver.CloseSocket
Unload Me
End Sub

Private Sub Receiver_KBPS(ByVal KBPS As Long)
lblKBPS.Caption = BytesToKB(KBPS) & " KB/Sec"
End Sub

Private Sub Receiver_ProgressUpdate(ByVal BytesReceived As Double, ByVal BytesTotal As Double)
On Error Resume Next
Bar.Value = BytesReceived
End Sub

Private Sub Receiver_ReceivedTransferRequest(ByVal FileTitle As String, ByVal FileSize As Double)
Bar.Max = FileSize
lblFN.Caption = FileTitle
lblFS.Caption = BytesToKB(FileSize)
End Sub

Private Sub Receiver_SocketError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
StatusBar.SimpleText = "Status: " & ErrorDescription
Receiver.CloseSocket
Receiver.Reset
End Sub

Private Sub Receiver_TransferComplete()
Bar.Value = Bar.Max
StatusBar.SimpleText = "Status: Download complete"
Receiver.Reset
If chkClose.Value Then Unload Me
End Sub

Private Sub Receiver_TransferStarted()
StatusBar.SimpleText = "Status: Downloading file..."
End Sub
