VERSION 5.00
Begin VB.Form frmAddAB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book - [Add Contact]"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
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
   Icon            =   "frmAddAB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1080
      Width           =   4215
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin RemoteFileManager.Button cmdAdd 
      Height          =   375
      Left            =   3000
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Add"
      ForeColor       =   -2147483630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Remote port:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hostname / IP address:"
      ForeColor       =   &H00BD662D&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmAddAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If Len(txtHost.Text) = 0 Then
    MsgBox "Enter a remote host or IP address", vbCritical, "Remote Host Required"
    txtHost.SetFocus
    Exit Sub
ElseIf Len(txtPort.Text) = 0 Then
    MsgBox "Enter a port to connect to", vbCritical, "Remote Port Required"
    txtPort.SetFocus
    Exit Sub
ElseIf Not IsNumeric(txtPort.Text) Then
    MsgBox "Enter a numeric value for the port", vbCritical, "Invalid Port Value"
    txtPort.SetFocus
    txtPort.SelStart = 0
    txtPort.SelLength = Len(txtPort.Text)
    Exit Sub
End If
Call AddressBookAdd(txtHost.Text, txtPort.Text, txtDesc.Text)
Call AddressBookLoadListView(frmAddressBook.LVAB)
Unload Me
End Sub

Private Sub txtPort_KeyPress(KeyAscii As Integer)
If Not IsNumeric(Chr$(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
