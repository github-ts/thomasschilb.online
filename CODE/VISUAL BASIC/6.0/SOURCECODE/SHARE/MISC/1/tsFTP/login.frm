VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Login"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3480
   LinkTopic       =   "Form3"
   ScaleHeight     =   2025
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbc 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cbo 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox tbp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   4
      Text            =   "tbp"
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox tbln 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   "tbln"
      Top             =   120
      Width           =   1935
   End
   Begin VB.CheckBox cxr 
      Caption         =   "&Remember my login info"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Login Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label3_Click()

End Sub

Private Sub cbc_Click()


Unload Me

End Sub

Private Sub cbo_Click()
loginCancel = False
loginName = tbln.Text
loginPassword = tbp.Text
loginRemember = cxr.Value

If loginRemember Then
    SaveSetting App.Title, "adv_ftp", "name", loginName
    SaveSetting App.Title, "adv_ftp", "pw", loginPassword
Else
    SaveSetting App.Title, "adv_ftp", "name", ""
    SaveSetting App.Title, "adv_ftp", "pw", ""
End If


Unload Me
End Sub

Private Sub Form_Load()

loginCancel = True
cxr.Value = 1
tbln.Text = GetSetting(App.Title, "adv_ftp", "name")
tbp.Text = GetSetting(App.Title, "adv_ftp", "pw")

    

End Sub

Private Sub Form_Resize()
tbln.SetFocus

If Len(tbln.Text) > 0 And Len(tbp.Text) > 0 Then cbo.SetFocus
End Sub
