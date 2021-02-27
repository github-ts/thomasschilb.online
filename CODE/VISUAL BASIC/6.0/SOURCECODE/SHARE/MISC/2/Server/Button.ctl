VERSION 5.00
Begin VB.UserControl Button 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2025
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
   ScaleHeight     =   3600
   ScaleWidth      =   2025
   ToolboxBitmap   =   "Button.ctx":0000
   Begin VB.Timer tmrChkStatus 
      Interval        =   150
      Left            =   840
      Top             =   1200
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      Height          =   300
      Left            =   0
      MouseIcon       =   "Button.ctx":0312
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   45
      Width           =   1935
   End
   Begin VB.Image imgUp 
      Height          =   330
      Left            =   0
      Picture         =   "Button.ctx":0464
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1905
   End
   Begin VB.Image imgOver 
      Height          =   330
      Left            =   0
      Picture         =   "Button.ctx":2AF9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1905
   End
   Begin VB.Image imgDown 
      Height          =   330
      Left            =   0
      Picture         =   "Button.ctx":55B9
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1905
   End
End
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private mbooButtonLighted As Boolean
Private mpoiCursorPos As POINTAPI

Public Event Click()

Public Event ButtonDown()

Public Event ButtonUp()

Private bIsActive As Boolean

Public Event MouseOver()



Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
If Button = 1 Then Call ShowDown
End Sub

Private Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Not Enabled Then Exit Sub
If Button = 1 Then
    Call ShowUp
    RaiseEvent Click
End If
End Sub

Private Sub tmrChkStatus_Timer()
If Not Enabled Then Exit Sub
Dim lonCStat As Long
Dim lonCurrhWnd As Long
tmrChkStatus.Enabled = False
lonCStat = GetCursorPos&(mpoiCursorPos)
lonCurrhWnd = WindowFromPoint(mpoiCursorPos.X, mpoiCursorPos.Y)
If mbooButtonLighted = False Then
    If lonCurrhWnd = UserControl.hWnd Then
    mbooButtonLighted = True
    Call ShowOver
    RaiseEvent MouseOver
    End If
Else
    If lonCurrhWnd <> UserControl.hWnd Then
    mbooButtonLighted = False
    Call ShowUp
    End If
End If
tmrChkStatus.Enabled = True
End Sub

Private Sub UserControl_InitProperties()
tmrChkStatus.Enabled = Ambient.UserMode
bIsActive = Ambient.UserMode
Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_Resize()
imgDown.Width = UserControl.Width
imgUp.Width = UserControl.Width
imgOver.Width = UserControl.Width
imgDown.Height = UserControl.Height
imgUp.Height = UserControl.Height
imgOver.Height = UserControl.Height
lblCaption.Width = UserControl.Width
lblCaption.Top = ((Height - lblCaption.Height) / 2) + 20
lblCaption.Left = (Width - lblCaption.Width) / 2
End Sub

Public Property Get Caption() As String
Caption = lblCaption.Caption
End Property

Public Property Let Caption(ByVal sNewValue As String)
lblCaption.Caption = sNewValue
UserControl.PropertyChanged "Caption"
End Property

Public Property Get FontName() As String
FontName = lblCaption.FontName
End Property

Public Property Let FontName(ByVal sNewValue As String)
lblCaption.FontName = sNewValue
UserControl.PropertyChanged "FontName"
End Property

Public Property Get FontSize() As Integer
FontSize = lblCaption.FontSize
End Property

Public Property Let FontSize(ByVal iNewValue As Integer)
lblCaption.FontSize = iNewValue
UserControl.PropertyChanged "FontSize"
End Property

Public Property Get Bold() As Boolean
Bold = lblCaption.FontBold
End Property

Public Property Let Bold(ByVal bNewValue As Boolean)
lblCaption.FontBold = bNewValue
UserControl.PropertyChanged "Bold"
End Property

Public Property Get Italic() As Boolean
Italic = lblCaption.FontItalic
End Property

Public Property Let Italic(ByVal bNewValue As Boolean)
lblCaption.FontItalic = bNewValue
UserControl.PropertyChanged "Italic"
End Property

Public Property Get Underline() As Boolean
Underline = lblCaption.FontUnderline
End Property

Public Property Let Underline(ByVal bNewValue As Boolean)
lblCaption.FontUnderline = bNewValue
UserControl.PropertyChanged "Underline"
End Property

Public Property Get ForeColor() As Long
ForeColor = lblCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal lNewValue As Long)
lblCaption.ForeColor = lNewValue
UserControl.PropertyChanged "lNewValue"
End Property

Private Sub ShowUp()
imgDown.Visible = False
imgOver.Visible = False
imgUp.Visible = True
RaiseEvent ButtonUp
End Sub

Private Sub ShowDown()
imgOver.Visible = False
imgUp.Visible = False
imgDown.Visible = True
RaiseEvent ButtonDown
End Sub

Private Sub ShowOver()
imgDown.Visible = False
imgUp.Visible = False
imgOver.Visible = True
RaiseEvent MouseOver
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", Caption, Ambient.DisplayName
    .WriteProperty "Bold", Bold, False
    .WriteProperty "Italic", Italic, False
    .WriteProperty "Underline", Underline, False
    .WriteProperty "FontName", FontName, "Tahoma"
    .WriteProperty "FontSize", FontSize, 8
    .WriteProperty "ForeColor", ForeColor, 0
    .WriteProperty "Enabled", Enabled, True
End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
tmrChkStatus.Enabled = Ambient.UserMode
With PropBag
    Caption = .ReadProperty("Caption", Ambient.DisplayName)
    Bold = .ReadProperty("Bold", False)
    Italic = .ReadProperty("Italic", False)
    Underline = .ReadProperty("Underline", False)
    FontName = .ReadProperty("FontName", "Tahoma")
    FontSize = .ReadProperty("FontSize", 8)
    ForeColor = .ReadProperty("ForeColor", 0)
    Enabled = .ReadProperty("Enabled", True)
End With
End Sub

Public Property Get Enabled() As Boolean
Enabled = lblCaption.Enabled
End Property

Public Property Let Enabled(ByVal bNewValue As Boolean)
lblCaption.Enabled = bNewValue
End Property
