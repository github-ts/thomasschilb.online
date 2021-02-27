VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Properties"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbc 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   20
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox tbl 
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox tbn 
      Height          =   285
      Left            =   1320
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   210
      Width           =   2415
   End
   Begin VB.CheckBox cxc 
      Caption         =   "&Compressed"
      Height          =   495
      Left            =   2520
      TabIndex        =   17
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox cxs 
      Caption         =   "&System"
      Height          =   495
      Left            =   2520
      TabIndex        =   16
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CheckBox cxh 
      Caption         =   "&Hidden"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox cxr 
      Caption         =   "&Read-only"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cbo 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   3960
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   3960
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   3960
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lt 
      AutoSize        =   -1  'True
      Caption         =   "lt"
      Height          =   195
      Left            =   1320
      TabIndex        =   15
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   405
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3960
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label la 
      AutoSize        =   -1  'True
      Caption         =   "la"
      Height          =   195
      Left            =   1320
      TabIndex        =   13
      Top             =   2880
      Width           =   120
   End
   Begin VB.Label lm 
      AutoSize        =   -1  'True
      Caption         =   "lm"
      Height          =   195
      Left            =   1320
      TabIndex        =   12
      Top             =   2520
      Width           =   150
   End
   Begin VB.Label lc 
      AutoSize        =   -1  'True
      Caption         =   "lc"
      Height          =   195
      Left            =   1320
      TabIndex        =   11
      Top             =   2160
      Width           =   120
   End
   Begin VB.Label ls 
      AutoSize        =   -1  'True
      Caption         =   "ls"
      Height          =   195
      Left            =   1320
      TabIndex        =   10
      Top             =   1560
      Width           =   105
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Attributes:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Accessed:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Modified:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Created:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label z3 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label z2 
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   660
   End
   Begin VB.Label z1 
      AutoSize        =   -1  'True
      Caption         =   "Name of file:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim f

Private Sub cbc_Click()
Unload Me
End Sub

Private Sub cbo_Click()

On Error GoTo Error

Dim ans As String
Dim attr As Long
If cxr.Value = 1 Then
    attr = attr Or ReadOnly
End If
If cxh.Value = 1 Then
    ans = MsgBox("If you make the file hidden, it may not be viewable. Continue?", vbYesNo, "Question")
    If ans = vbYes Then attr = attr Or Hidden
End If
If cxs.Value = 1 Then
    ans = MsgBox("If you make the file a system file, it may not be viewable. Continue?", vbYesNo, "Question")
    If ans = vbYes Then attr = attr Or System
End If
If cxc.Value = 1 Then
    attr = attr Or Compressed
End If
f.Attributes = attr

If f.Name <> tbn.Text Then f.Name = tbn.Text
Unload Me
Exit Sub

Error:
    MsgBox ERR.Description, vbCritical, "Error"
    Unload Me

End Sub

Private Sub Form_Load()
'Dim fr As Folder
'fr = fs.GetFolder
On Error Resume Next

If propertyType = WANT_FILE_PROPERTY Then
    Set f = fs.GetFile(fileForProperty)
    Me.Caption = "File Properties"
Else
    Set f = fs.GetFolder(fileForProperty)
    Me.Caption = "Folder Properties"
End If
tbn.Text = f.Name
lt.Caption = f.Type
tbl.Text = fileForProperty
ls.Caption = Format(f.Size / 1024#, "#.##") & " KB (" & f.Size & " bytes)"
lc.Caption = f.DateCreated
lm.Caption = f.DateLastModified
la.Caption = f.DateLastAccessed
If (f.Attributes And ReadOnly) <> 0 Then
    cxr.Value = 1
Else
    cxr.Value = 0
End If
If (f.Attributes And Hidden) <> 0 Then
'MsgBox "attr: " & f.Attributes & " hidden: " & Hidden
    cxh.Value = 1
Else
    cxh.Value = 0
End If

If (f.Attributes And System) <> 0 Then
    cxs.Value = 1
Else
    cxs.Value = 0
End If

If (f.Attributes And Compressed) <> 0 Then
    cxc.Value = 1
Else
    cxc.Value = 0
End If



End Sub

Private Sub Form_Resize()

Line4.Y1 = Line1.Y1 + 25
Line4.Y2 = Line1.Y2 + 25

Line5.Y1 = Line2.Y1 + 25
Line5.Y2 = Line2.Y2 + 25

Line6.Y1 = Line3.Y1 + 25
Line6.Y2 = Line3.Y2 + 25

End Sub

Private Sub tbl_KeyPress(KeyAscii As Integer)
'eat the key no matter what user enters
KeyAscii = 0
End Sub
