VERSION 5.00
Begin VB.Form frmFileInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Information"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
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
   Icon            =   "frmFileInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RemoteFileManager.Button cmdOk 
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Ok"
      ForeColor       =   -2147483630
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Remote File Manager - File Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   10
      Top             =   120
      Width           =   3315
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   0
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      BorderWidth     =   3
      X1              =   0
      X2              =   5760
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File location:"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lblFSK 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File size (KB):"
      Height          =   195
      Left            =   720
      TabIndex        =   5
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label lblFSB 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File size (Bytes):"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   360
      Picture         =   "frmFileInfo.frx":058A
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File name:"
      Height          =   195
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label lblFN 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "File information:"
      Height          =   195
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   5760
      X2              =   960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmFileInfo.frx":0B14
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmFileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload Me
End Sub
