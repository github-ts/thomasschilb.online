VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.UserControl FileSender 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSWinsockLib.Winsock sckData 
      Left            =   1440
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6704
   End
   Begin MSWinsockLib.Winsock sckMain 
      Left            =   360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   4509
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   0
      Picture         =   "FileSender.ctx":0000
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "FileSender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetInputState Lib "user32" () As Long

Public Event Connected()
Public Event SocketError(ByVal ErrorNumber As Long, ByVal ErrorDescription As String)
Public Event SocketClosed()
Public Event TransferStopped()
Public Event TransferStarted()
Public Event TransferComplete()
Public Event DataSent(DataLen As Long)
Public Event DataReceived(DataLen As Long)

Const DELIM As String = "_"
Const EOP As String = "++"

Dim lPackSize As Long

Dim sPath As String
Dim sTitle As String

Dim bCon As Boolean

Dim CurByte As Long
Dim dSendTotal As Long
Dim dFileSize As Long

Public Property Get PacketSize() As Long
PacketSize = lPackSize
End Property

Public Property Let PacketSize(ByVal lNewValue As Long)
lPackSize = lNewValue
End Property

Public Property Get FilePath() As String
FilePath = sPath
End Property

Public Property Let FilePath(ByVal sNewValue As String)
sPath = sNewValue
End Property

Public Property Get FileTitle() As String
FileTitle = sTitle
End Property

Public Property Let FileTitle(ByVal sNewValue As String)
sTitle = sNewValue
End Property

Private Sub sckData_ConnectionRequest(ByVal requestID As Long)
sckData.Close
sckData.Accept requestID
RaiseEvent TransferStarted
Call SendFile(FilePath)
End Sub

Private Sub sckData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent SocketError(Number, Description)
'Call CloseSocket
'bCon = False
End Sub

Private Sub sckData_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
dSendTotal = dSendTotal + bytesSent
RaiseEvent DataSent(bytesSent)
If dSendTotal >= dFileSize Then
    Call SendTransferComplete
    'Call Reset
End If
End Sub

Private Sub sckMain_Close()
'Call CloseSocket
'bCon = False
RaiseEvent SocketClosed
End Sub

Private Sub sckMain_Connect()
bCon = True
RaiseEvent Connected
Call SendFileInfo
End Sub

Private Sub sckMain_DataArrival(ByVal bytesTotal As Long)
Dim sData As String, sBuff() As String
Dim iLoop As Integer
sckMain.GetData sData
RaiseEvent DataReceived(Len(sData))
If InStr(1, sData, EOP) = 0 Then Exit Sub
sBuff() = Split(sData, EOP)
For iLoop = 0 To UBound(sBuff)
    If Len(sBuff(iLoop)) > 0 Then
        If Left(sBuff(iLoop), 3) = "NET" Then
            Call ParseTransferComplete(sBuff(iLoop))
        ElseIf Left(sBuff(iLoop), 3) = "STD" Then
            Call ParseTransferStopped(sBuff(iLoop))
        ElseIf Left(sBuff(iLoop), 3) = "STP" Then
            Call ParseStopTransfer(sBuff(iLoop))
        End If
    End If
Next iLoop
End Sub

Private Sub sckMain_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
RaiseEvent SocketError(Number, Description)
'Call CloseSocket
'bCon = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    PacketSize = .ReadProperty("PacketSize", 4096)
    FilePath = .ReadProperty("FilePath", Empty)
    FileTitle = .ReadProperty("FileTitle", "(File title)")
    RemoteHost = .ReadProperty("RemoteHost", Empty)
    MainPort = .ReadProperty("MainPort", 4509)
    TransferPort = .ReadProperty("TransferPort", 6704)
End With
End Sub

Private Sub UserControl_Resize()
Width = imgIcon.Width
Height = imgIcon.Height
End Sub

Public Property Get RemoteHost() As String
RemoteHost = sckMain.RemoteHost
End Property

Public Property Let RemoteHost(ByVal sNewValue As String)
sckMain.RemoteHost = sNewValue
End Property

Public Property Get MainPort() As Integer
MainPort = sckMain.RemotePort
End Property

Public Property Let MainPort(ByVal iNewValue As Integer)
sckMain.RemotePort = iNewValue
End Property

Public Property Get TransferPort() As Integer
TransferPort = sckData.LocalPort
End Property

Public Property Let TransferPort(ByVal iNewValue As Integer)
sckData.LocalPort = iNewValue
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "PacketSize", PacketSize, 4096
    .WriteProperty "FilePath", FilePath, Empty
    .WriteProperty "FileTitle", FileTitle, "(File Title)"
    .WriteProperty "RemoteHost", RemoteHost, Empty
    .WriteProperty "MainPort", MainPort, 4509
    .WriteProperty "TransferPort", TransferPort, 6704
End With
End Sub

Public Sub CloseSocket()
sckMain.Close
sckData.Close
bCon = False
Call Reset
End Sub

Public Sub Reset()
CurByte = 0
dSendTotal = 0
dFileSize = 0
End Sub

Public Sub Connect()
Call CloseSocket
bCon = False
'Call Reset
With sckMain
    .RemoteHost = RemoteHost
    .RemotePort = MainPort
    .Connect
End With
End Sub

Private Sub SendFileInfo()
Dim sPacket As String
dFileSize = FileLen(FilePath)
If Not bCon Then Exit Sub
sckData.Close
sckData.LocalPort = TransferPort
sckData.Listen
sPacket = "SND" & DELIM & FileTitle & DELIM & dFileSize
sckMain.SendData sPacket & EOP
End Sub

Private Sub SendTransferComplete()
Dim sPacket As String
If Not bCon Then Exit Sub
sPacket = "NET" & DELIM & "Sent"
sckMain.SendData sPacket & EOP
End Sub

Private Sub SendFile(sFilePath As String)
Dim FF As Integer: FF = FreeFile
Dim B As Long
Dim bBuffer() As Byte
Open sFilePath For Binary Access Read As #FF
ReDim bBuffer(1 To 4096) As Byte
Do Until (dFileSize - CurByte) < 4096
    Get #FF, CurByte + 1, bBuffer()
    CurByte = CurByte + 4096
    On Error GoTo Err
    sckData.SendData bBuffer
    If GetInputState Then DoEvents
Loop

Dim PrevPackSize As Long
PrevPackSize = dFileSize - CurByte
ReDim bBuffer(1 To PrevPackSize) As Byte
Get #FF, CurByte + 1, bBuffer()
CurByte = CurByte + PrevPackSize
sckData.SendData bBuffer
Close #FF
Exit Sub
Err:
Exit Sub
End Sub

Private Sub ParseTransferComplete(ByVal Data As String)
'NET/Sent
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Sent" Then
    RaiseEvent TransferComplete
    'Call CloseSocket
    'bCon = False
End If
End Sub

Private Sub SendStopTransfer()
'STP/"Stop"
Dim sPacket As String
sPacket = "STP" & DELIM & "Stop"
If Not bCon Then Exit Sub
sckMain.SendData sPacket
End Sub

Private Sub SendTransferStopped()
'SDT/Stopped
Dim sPacket As String
sPacket = "SDT" & DELIM & "Stopped"
If Not bCon Then Exit Sub
sckMain.SendData sPacket
End Sub

Private Sub ParseTransferStopped(ByVal Data As String)
'SDT/"Stopped"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Stopped" Then
    'Call CloseSocket
    'bCon = False
    RaiseEvent TransferStopped
End If
End Sub

Private Sub ParseStopTransfer(ByVal Data As String)
'STP"/"Stop"
Dim sBuff() As String: sBuff() = Split(Data, DELIM)
If sBuff(1) = "Stop" Then
    Call SendTransferStopped
    RaiseEvent TransferStopped
End If
End Sub

Public Sub StopTransfer()
Call SendStopTransfer
End Sub
