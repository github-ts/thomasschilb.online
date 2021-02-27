VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmAddressBook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Address Book"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6120
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
   Icon            =   "frmAddressBook.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Added contacts "
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   5655
      Begin MSComctlLib.ImageList ILAB 
         Left            =   1320
         Top             =   960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddressBook.frx":3482
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddressBook.frx":3E94
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAddressBook.frx":7326
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LVAB 
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ILAB"
         SmallIcons      =   "ILAB"
         ColHdrIcons     =   "ILAB"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IP Address"
            Object.Width           =   2540
            ImageIndex      =   1
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Port"
            Object.Width           =   2540
            ImageIndex      =   2
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   4304
            ImageIndex      =   3
         EndProperty
      End
   End
   Begin RemoteFileManager.Button cmdAdd 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Add..."
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdRemove 
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Remove"
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdRemoveAll 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Remove All"
      ForeColor       =   -2147483630
   End
   Begin RemoteFileManager.Button cmdConnect 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "Connect to..."
      ForeColor       =   -2147483630
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Remote File Manager® remembers your contacts so you don't have to!"
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
      Height          =   555
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   4680
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmAddressBook.frx":A7B8
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmAddressBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
frmAddAB.Show vbModal
End Sub

Private Sub cmdConnect_Click()
If LVAB.ListItems.Count = 0 Then Exit Sub
Dim cSel As New Collection
Set cSel = GetSelectedContactInfo
If cSel.Count = 0 Then
    MsgBox "Select a contact to connect to", vbCritical, "Contact Required"
    Exit Sub
ElseIf Len(cSel.Item(1)) = 0 Then
    Exit Sub
Else
    With frmMain
        .txtHost.Text = cSel.Item(1)
        .txtPort.Text = cSel.Item(2)
        .cmdConnect_Click
    End With
    Unload Me
End If
End Sub

Private Sub cmdRemove_Click()
If LVAB.ListItems.Count = 0 Then Exit Sub
Dim cSel As New Collection
Set cSel = GetSelectedContactInfo
If cSel.Count = 0 Then
    MsgBox "Select a contact to remove", vbCritical, "Contact Required"
    Exit Sub
ElseIf Len(cSel.Item(1)) = 0 Then
    Exit Sub
Else
    Call AddressBookRemove(cSel.Item(1), cSel.Item(2))
    ' Call RemoveListViewItem(LVAB, cSel.Item(1))
    Call AddressBookLoadListView(LVAB)
End If
End Sub

Private Sub cmdRemoveAll_Click()
If LVAB.ListItems.Count = 0 Then Exit Sub
Dim vbRep As VbMsgBoxResult
vbRep = MsgBox("Are you sure you want to remove all contacts?", vbQuestion + vbYesNo, "Remove All Contacts")
If vbRep = vbYes Then
    Call KillFile(App.Path & ADB_FILE)
    LVAB.ListItems.Clear
End If
End Sub

Private Sub Form_Load()
Call AddressBookLoadListView(LVAB)
End Sub

Function GetSelectedContactInfo() As Collection
Dim cRet As New Collection
On Error Resume Next
Dim lSel As Long, sSel As String
With LVAB
    lSel = .SelectedItem.Index
    If lSel = 0 Then Exit Function
    sSel = Trim$(.ListItems(lSel).Text)
    cRet.Add sSel, sSel
    sSel = Trim$(.ListItems(lSel).ListSubItems(1).Text)
    cRet.Add sSel, sSel
    sSel = Trim$(.ListItems(lSel).ListSubItems(2).Text)
    cRet.Add sSel, sSel
End With
Set GetSelectedContactInfo = cRet
End Function
