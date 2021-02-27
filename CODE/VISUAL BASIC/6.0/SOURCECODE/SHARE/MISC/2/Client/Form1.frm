VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Transter Client"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsend 
      Left            =   3000
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComDlg.CommonDialog coFile 
      Left            =   3240
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   255
      Left            =   5040
      TabIndex        =   10
      Top             =   1560
      Width           =   255
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5415
      TabIndex        =   3
      Top             =   2460
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "Send"
         Height          =   375
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "203.91.138.33"
      Top             =   840
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   5760
      TabIndex        =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock wLogin 
      Left            =   4560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   5760
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock wServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   885
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public accok As Boolean
Dim FileContent As String
Dim Buf() As Byte
Dim bufPos As Long
Dim Sendbytes As Long

Private Sub Command1_Click()
If Len(Trim(Text2.Text)) = 0 Then
Text2.SetFocus
Exit Sub
End If
Winsock1.Close
Winsock1.RemoteHost = txtServer.Text
Winsock1.LocalPort = 11111
Winsock1.RemotePort = 7777
Winsock1.Bind Winsock1.LocalPort
Winsock1.SendData "IP=" & Winsock1.LocalIP
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
coFile.CancelError = False
coFile.ShowOpen
txtFile.Text = coFile.FileName
End Sub

Private Sub Form_Load()
wServer.Close
wServer.Bind 7778, wServer.LocalIP
wLogin.Close
wLogin.Bind 7779, wServer.LocalIP
End Sub

Private Sub winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Winsock1.GetData gotstring, vbString
End Sub

Private Sub wsend_Connect()
Winsock1.SendData Buf
End Sub

Private Sub wServer_DataArrival(ByVal bytesTotal As Long)
wServer.GetData gotstring, vbString
Winsock1.SendData "Password=" & Text2.Text
Select Case gotstring
Case "Accept"
SendFile txtFile
Case "OK All"
accok = True
End Select
Select Case gotstring
Case "Yes"
Winsock1.SendData "Ready?"
Case "OK"
If accok Then
If Len(txtFile.Text) > 0 Then
Winsock1.SendData "File=" & txtFile.Text
End If
End If
Case "Reject"
MsgBox "Failed to connect to server", vbOKOnly + vbInformation, "Client":
End
End Select
End Sub

Sub SendFile(strFile As String)
On Error Resume Next
txtsource = strFile
FileContent = Space(FileLen(txtsource))
Open txtsource For Binary Access Read As #1
Get #1, , FileContent
Close #1
With wsend
  .RemoteHost = txtServer.Text
  .RemotePort = 1215
  .SendData FileContent
End With
End Sub



