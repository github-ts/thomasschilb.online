VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Same File Exists"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton obda 
      Caption         =   "Di&sgard all"
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   3120
      Width           =   3255
   End
   Begin VB.OptionButton obdi 
      Caption         =   "D&isgard"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox tbf 
      Height          =   285
      Left            =   2160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2280
      Width           =   2655
   End
   Begin VB.OptionButton obd 
      Caption         =   "&Do not overwrite"
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2160
      Width           =   3615
   End
   Begin VB.OptionButton oboa 
      Caption         =   "O&verwrite all"
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.OptionButton obo 
      Caption         =   "&Overwrite"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   3615
   End
   Begin VB.CommandButton cbo 
      Caption         =   "O&K"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label l 
      AutoSize        =   -1  'True
      Caption         =   "A file by the same name is found. What do you want to do?"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4170
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this form is invoked when user wants to save file in a location already occupied
'in that case, ask user whether he wants to overwrite or do not overwrite


Option Explicit

Private Sub cbo_Click()
'MsgBox "its: " & tbf.Text
If obd.Value = True Then
    If Mid(myFilePath, InStrRev(myFilePath, "\") + 1) = _
        tbf.Text Then
        MsgBox "You must enter a new name if you don't wanna overwrite."
        Exit Sub
    ElseIf Not is_filename_valid(tbf.Text) Then
        MsgBox tbf.Text & " is not a valid file name."
        Exit Sub
    End If
End If
newFilePath = Left(myFilePath, InStrRev(myFilePath, "\")) & tbf.Text
If obo.Value = True Then
    fileFoundOption = OVERWRITE
ElseIf oboa.Value = True Then
    fileFoundOption = OVERWRITE_ALL
ElseIf obd.Value = True Then
    fileFoundOption = DO_NOT_OVERWRITE
ElseIf obdi.Value = True Then
    fileFoundOption = DISCARD
Else
    fileFoundOption = DISCARD_ALL
End If
Unload Me
End Sub

Private Sub Form_Load()

l.Caption = "A file by the same name is found:" & vbCrLf & _
    myFilePath & vbCrLf & "What do you want to do?"
tbf.Text = Mid(myFilePath, InStrRev(myFilePath, "\") + 1)
obo.Value = True
tbf.Enabled = False

End Sub

Private Sub obd_Click()
tbf.Enabled = True
End Sub



Private Sub obo_Click()
tbf.Enabled = False
End Sub

Private Sub oboa_Click()
tbf.Enabled = False
End Sub
