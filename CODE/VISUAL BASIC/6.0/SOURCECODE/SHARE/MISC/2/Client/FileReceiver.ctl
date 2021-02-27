VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl FileReceiver 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Timer tmrKBPS 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   1560
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   0
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   4509
   End
   Begin MSWinsockLib.Winsock sckData 
      Left            =   480
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6704
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "FileReceiver.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FileReceiver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Event ClientConnected(ByVal IPAddress As String)
Public Event TransferStarted()
Public Event TransferStopped()
Public Event TransferComplete()
Public Event SocketError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event SocketClosed()
Public Event ReceivedTransferRequest(ByVal FileTitle As String, ByVal FileSize As Double)
Public Event ProgressUpdate(ByVal BytesReceived As Double, ByVal BytesTotal As Double)
Public Event KBPS(ByVal KBPS As Long)

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const DELIM As String = "_"
Private Const EOP As String = "++"

Dim sOutDir As String
Dim sFileName As String

Dim bCon As Boolean

Dim dFileSize As Double
Dim dBytesRec As Long

Dim lDownloadSpeed As Long
Dim lDownloadSecond As Long

Dim iFN As Integer

Private Sub sckData_Connect()
bCon = True
tmrKBPS.Enabled = True
Dim sTmpDir As String
sTmpDir = ReceiveDirectory
If Not Right(sTmpDir, 1) = "\" Then sTmpDir = sTmpDir & "\"
Open sTmpDir & sFileName For Binary Access Write As #iFN
RaiseEvent TransferStarted
End Sub

Private Sub sckData_DataArrival(ByVal BytesTotal As Long)
Dim sData As String
sckData.GetData sData
dBytesRec = dBytesRec + BytesTotal
RaiseEvent ProgressUpdate(dBytesRec, dFileSize)
Put #iFN, , sData
End Sub

Private Sub sckData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Close #iFN
RaiseEvent SocketError(Number, Description)
End Sub

Private Sub sckMain_Close()
RaiseEvent SocketClosed
'Call CloseSocket
End Sub

Private Sub sckMain_ConnectionRequest(ByVal requestID As Long)
sckMain.Close
sckMain.Accept requestID
RaiseEvent ClientConnected(sckMain.RemoteHostIP)
End Sub

Private Sub sckMain_DataArrival(ByVal BytesTotal As Long)
Dim sData As String, sBuff() As String, iLoop As Integer, sTmp As String
sckMain.GetData sData
If Not InStr(1, sData, EOP) > 0 Then Exit Sub
sBuff() = Split(sData, EOP)
For iLoop = 0 To UBound(sBuff)
    sTmp = sBuff(iLoop)
    If Len(sTmp) > 0 Then
        If Left(sTmp, 3) = "SND" Then
            Call ParseFileRequest(sTmp)
        ElseIf Left(sTmp, 3) = "NET" Then
            Call ParseTransferComplete(sTmp)
        ElseIf Left(sTmp, 3) = "STP" Then
            Call ParseStopTransfer(sTmp)
        ElseIf Left(sTmp, 3) = "SDT" Then
            Call ParseTransferStopped(sTmp)
        End If
    End If
Next iLoop
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent SocketError(Number, Description)
'Call CloseSocket
End Sub

Private Sub tmrKBPS_Timer()
On Error Resume Next
lDownloadSpeed = dBytesRec - lDownloadSecond
lDownloadSecond = dBytesRec
Dim lKBPS As Long
lKBPS = lDownloadSpeed
RaiseEvent KBPS(lKBPS)
End Sub

Private Sub UserControl_Initialize()
iFN = FreeFile
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    MainPort = .ReadProperty("MainPort", 4509)
    TransferPort = .ReadProperty("TransferPort", 6704)
    ReceiveDirectory = .ReadProperty("ReceiveDirectory", TempDir)
End With
End Sub

Private Sub UserControl_Resize()
Height = imgIcon.Height
Width = imgIcon.Width
End Sub

Public Property Get MainPort() As Integer
MainPort = sckMain.LocalPort
End Property

Public Property Let MainPort(ByVal iNewValue As Integer)
sckMain.LocalPort = iNewValue
End Property

Public Property Get TransferPort() As Integer
TransferPort = sckData.RemotePort
End Property

Public Property Let TransferPort(ByVal iNewValue As Integer)
sckData.RemotePort = iNewValue
End Property

Public Property Get ReceiveDirectory() As String
ReceiveDirectory = sOutDir
End Property

Public Property Let ReceiveDirectory(ByVal sNewValue As String)
sOutDir = sNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "MainPort", MainPort, 4509
    .WriteProperty "TransferPort", TransferPort, 6704
    .WriteProperty "ReceiveDirectory", ReceiveDirectory, TempDir
End With
End Sub

Private Function TempDir() As String
On Error Resume Next
Dim sBuff As String * 255
Dim lRet As Long, sRet As String
lRet = GetTempPath(sBuff, 255)
sRet = Trim$(sBuff)
If Not Right(sRet, 1) = "\" Then sRet = sRet & "\"
TempDir = sRet
End Function

Public Sub Reset()
Close #iFN
dFileSize = 0
sFileName = Empty
dBytesRec = 0
End Sub

Public Sub CloseSocket()
Close #iFN
sckMain.Close
sckData.Close
bCon = False
Call Reset
End Sub

Public Sub Listen()
Call CloseSocket
sckMain.LocalPort = MainPort
sckMain.Listen
End Sub

Private Sub ParseFileRequest(ByVal Data As String)
'SND/File title/File size
Call Reset
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
sFileName = sBuff(1)
dFileSize = CDbl(sBuff(2))
RaiseEvent ReceivedTransferRequest(sFileName, dFileSize)
With sckData
    .Close
    .RemoteHost = sckMain.RemoteHostIP
    .RemotePort = TransferPort
    .Connect
End With
End Sub

Private Sub ParseTransferComplete(ByVal Data As String)
'NET/"Sent"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Sent" Then
    Close #iFN
    Call Reset
    sckData.Close
    RaiseEvent TransferComplete
    tmrKBPS.Enabled = False
    Call SendTransferComplete
End If
End Sub

Private Sub SendTransferComplete()
Dim sPacket As String
If Not bCon Then Exit Sub
sPacket = "NET" & DELIM & "Sent"
sckMain.SendData sPacket & EOP
End Sub

Private Sub SendStopTransfer()
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "STP" & DELIM & "Stop"
sckMain.SendData sPacket & EOP
End Sub

Private Sub SendTransferStopped()
If Not bCon Then Exit Sub
Dim sPacket As String
sPacket = "SDT" & DELIM & "Stopped"
sckMain.SendData sPacket & EOP
End Sub

Private Sub ParseTransferStopped(ByVal Data As String)
'SDT/Stopped
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Stopped" Then
    'Call CloseSocket
    'Call Reset
    RaiseEvent TransferStopped
End If
End Sub

Private Sub ParseStopTransfer(ByVal Data As String)
'STP/Stop
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Stop" Then
    Call Reset
    Call SendTransferStopped
    RaiseEvent TransferStopped
End If
End Sub

Public Sub StopTransfer()
Call SendStopTransfer
End Sub

