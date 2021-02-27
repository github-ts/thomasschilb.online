VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Options"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   LinkTopic       =   "Form2"
   ScaleHeight     =   2505
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tbn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Text            =   "tbn"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.CommandButton cbc 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cbo 
      Caption         =   "OK"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
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
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   "tbp"
      Top             =   240
      Width           =   1215
   End
   Begin VB.OptionButton obdwr2 
      Caption         =   "Dir without recursion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.OptionButton obdwr 
      Caption         =   "Dir with recursion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.OptionButton obf 
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label ln 
      AutoSize        =   -1  'True
      Caption         =   "Save under "
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
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbc_Click()


Unload Me

End Sub

Private Sub cbo_Click()

If obf.Value Then
    tranOption = TRAN_FILE
    defaultFile = tbn.Text
ElseIf obdwr.Value Then
    tranOption = TRAN_DIR_R
    defaultDir = tbn.Text
Else
    tranOption = TRAN_DIR
    defaultDir = tbn.Text
End If

pattern = tbp.Text

Unload Me

End Sub

Private Sub Form_Load()

tranOption = TRAN_CANCEL
obf.Value = True
tbp.Text = pattern
If defaultFile = "" Then
    obf.Enabled = False
    obdwr.Value = True
End If

End Sub


Private Sub obdwr_Click()
ln.Caption = "Dir Name:"
tbn.Text = defaultDir
tbp.Enabled = True
Label1.Enabled = True
End Sub

Private Sub obdwr2_Click()
ln.Caption = "Dir Name:"
tbn.Text = defaultDir
tbp.Enabled = True
Label1.Enabled = True
End Sub

Private Sub obf_Click()
ln.Caption = "File Name:"
tbn.Text = defaultFile
tbp.Enabled = False
Label1.Enabled = False
End Sub
