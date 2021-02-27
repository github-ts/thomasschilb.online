VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "tsFTP"
   ClientHeight    =   7950
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9135
   Icon            =   "tsFTP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cbf 
      Caption         =   "Refresh &Files"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   6600
      Width           =   2535
   End
   Begin VB.CommandButton cbr 
      Caption         =   "Refresh &Dir"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2640
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cbs 
      Caption         =   "Terminate (cbs)"
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
      Left            =   6000
      TabIndex        =   21
      Top             =   600
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   120
      TabIndex        =   19
      Top             =   3840
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2565
      Left            =   120
      TabIndex        =   18
      Top             =   960
      Width           =   2535
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   2535
   End
   Begin VB.Timer speedTimer 
      Left            =   2760
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar pbs 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.TextBox tbl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   5880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      Text            =   "tsFTP.frx":058A
      Top             =   600
      Width           =   3135
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   7560
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lbfi 
      Height          =   2985
      ItemData        =   "tsFTP.frx":0590
      Left            =   3240
      List            =   "tsFTP.frx":0597
      TabIndex        =   12
      Top             =   3840
      Width           =   2535
   End
   Begin VB.ListBox lbf 
      Height          =   2790
      ItemData        =   "tsFTP.frx":05A1
      Left            =   3240
      List            =   "tsFTP.frx":05A8
      TabIndex        =   11
      Top             =   960
      Width           =   2535
   End
   Begin VB.ComboBox cbsd 
      Height          =   315
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Width           =   2535
   End
   Begin VB.CommandButton cbput 
      Caption         =   ">>"
      Height          =   735
      Left            =   2760
      TabIndex        =   9
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton cbget 
      Caption         =   "<<"
      Height          =   615
      Left            =   2760
      TabIndex        =   8
      Top             =   3840
      Width           =   375
   End
   Begin VB.CommandButton cbt 
      Caption         =   "Log out (cbt)"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cbc 
      Caption         =   "Log in (cbc)"
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
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1935
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
      Height          =   345
      Left            =   3960
      TabIndex        =   3
      Text            =   "tbp"
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox tbip 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   2
      Text            =   "tbip"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   9120
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9120
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label ls 
      AutoSize        =   -1  'True
      Caption         =   "status"
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
      TabIndex        =   20
      Top             =   600
      Width           =   525
   End
   Begin VB.Label lspeed 
      Caption         =   "lspeed"
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
      Left            =   120
      TabIndex        =   16
      Top             =   7440
      Width           =   3015
   End
   Begin VB.Label lbytes 
      Caption         =   "lbytes"
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
      Left            =   3240
      TabIndex        =   14
      Top             =   7440
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Server's Files:"
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
      Left            =   3240
      TabIndex        =   7
      Top             =   240
      Width           =   1275
   End
   Begin VB.Label Label3 
      Caption         =   "My Files:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label lp 
      Caption         =   "Port:"
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
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Label li 
      Caption         =   "IP:"
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
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu mmain 
      Caption         =   "&Main"
      Begin VB.Menu mexport 
         Caption         =   "E&xport Log"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnone6 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mfolder 
      Caption         =   "&Folder"
      Begin VB.Menu mopen2 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnone4 
         Caption         =   "-"
      End
      Begin VB.Menu mrename2 
         Caption         =   "&Rename"
      End
      Begin VB.Menu mdelete2 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnone5 
         Caption         =   "-"
      End
      Begin VB.Menu mproperties2 
         Caption         =   "&Properties"
      End
   End
   Begin VB.Menu mfile 
      Caption         =   "F&ile"
      Begin VB.Menu mopen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnone 
         Caption         =   "-"
      End
      Begin VB.Menu mrename 
         Caption         =   "&Rename"
         Shortcut        =   ^R
      End
      Begin VB.Menu mdelete 
         Caption         =   "&Delete"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnone2 
         Caption         =   "-"
      End
      Begin VB.Menu mproperties 
         Caption         =   "&Properties"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mabout2 
      Caption         =   "&About"
      Begin VB.Menu mabout 
         Caption         =   "A&bout"
         Shortcut        =   ^B
      End
      Begin VB.Menu mhelp 
         Caption         =   "&Help"
         Shortcut        =   ^H
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this is FTP client that talk to FTP server in protocols defined by me

'when server gets a connection request from client, the following happens:
'server sends PP to client, client pops up a login window for user to enter login info
'if user presses Cancel, client sends PQ and server closes connection
'if user presses Ok, client sends PGname&passowrd and server checks its account list
'if any error occurs, server sends PXerror and closes connection
'otherwise, server sends PY. then both sides become CONN

'to get drive list, do
'mystate = want_drives  client sends PD; server sends "msg_length&A:&C:&D:&E:"
'to get folder list, do
'mystate = want_folders  client sends PF; server sends "msg_length&folder1&folder2&..."
'to get file list, do
'mystate = want_files  client sends PI; server sends "msg_length&file1&file2&..."

'to get a file from server, do
'mystate = REQ_FILE; client sends PLftpFilePath; server sends PBftpFileSize
'server's state is SEND_FILE, client sends PAchunkSize
'server sends next chunk, client sends PAchunkSize
'goes on until server sends last packet, then client will send last PA
'then server's state goes to CONN, and so does client's
'note that server's state doesn't go to CONN until it receives the last PA

'to get an entire dir from server, do
'mystate = REQ_DIR; client sends PZ0ftpFolderPath&pattern or PZ1ftpFolderPath&pattern
'depending on if user wants recursive folder or non-recursive
'server sends PBftpFilePath&ftpFileSize and changes state to SEND_DIR
'client's state is RECV_DIR; client sends PA0
'server sends next chunk...
'goes on until the last packet is sent; client's state goes to REQ_DIR
'server updates filePointer, sends next file in line, sends next
'PBfitpFilePath&ftpFileSize to client
'client's state is RECV_DIR, sends PA0...
'goes on until the last packet is sent; client's state goes to REQ_DIR
'server updates filePointer but it found all files are sent, it sends PE and
'its state goes to CONN
'client got PE and learns it's got everything; its state is CONN

'to upload a file to server, do
'mystate=ATTEMPT_SEND, client sends PBftpFilePath&ftpFileSize
'sever's state = RECV_FILE, it sends PA0
'client sends next chunk, server sends PA<bytesTotal>
'goes on until last packet is sent
'both' states go to CONN

'to upload a folder to server, I use a different approach from downloading a folder
'first client collects all files it wants to upload in a string array
'then i use a for loop to send it one by one; in the for loop, i do
'PBftpFilePath&ftpFileSize, then mystate becomes ATTEMPT_SEND
'then doEvents until mystate becomes CONN, meaning that file is done uploading
'of course i need to do some error checking..
'brilliant isn't it?

'note that winsock could chop ur msg in multiple chunks, so i need to take care of that
'i only handle data when msg length is equal to the given msg length

'command has no delimiters!!!! but replies generally do


Option Explicit

'if path doesn't exist this sub does nothing
Public Sub open_file_or_dir(path As String)

ShellExecute Me.hwnd, "open", path, _
        vbNullString, vbNullString, SW_SHOWNORMAL

End Sub

Public Function socket_send(s) As Boolean

If ws.STATE = sckConnected Then
    Call ws.SendData(s)
    socket_send = True
Else
    update_log "I want to send the following: " & vbCrLf & s & _
        ", but socket state is " & get_socket_state(ws.STATE)
    socket_send = False
End If

End Function


Public Function recv_chunk(bb() As Byte, bb_len As Long) As Boolean

'MsgBox ("in recv_file")

If (currFilePos <= 0) Then currFilePos = 1

Dim fh As Long
fh = FreeFile
If myFilePath = "" Or Len(myFilePath) < 1 Then
    MsgBox ("Impossible error in recv_chunk: myFilepath is invalid")
    Exit Function
End If






Open myFilePath For Binary As fh
Seek fh, LOF(fh) + 1
    
If (currFilePos + bb_len - 1 < ftpFileSize) Then
    currFilePos = currFilePos + bb_len
    recv_chunk = False
    
ElseIf currFilePos + bb_len - 1 = ftpFileSize Then
    recv_chunk = True

Else
    Call MsgBox("ftpFileSize: " & ftpFileSize & " currFilePos: " & currFilePos & _
        " bb_len: " & bb_len, vbCritical, "Error")
    recv_chunk = True
    Close #fh
    Exit Function
End If
    
Put fh, , bb
Close #fh


'Call update_buf(MODE.STATE, bb_len & " bytes received")
pbs.Value = currFilePos - 1
lbytes.Caption = (currFilePos - 1) & " / " & ftpFileSize & " bytes"

End Function

Public Function get_next_chunk(ByRef bb() As Byte, ByRef bb_len As Long) As Boolean
   
'MsgBox "in get chunk: " & myFilePath & " currFilePos: " & currFilePos
   
If (currFilePos <= 0) Then currFilePos = 1

Dim fh As Long
fh = FreeFile
Open myFilePath For Binary As #fh
Seek fh, currFilePos

myChunkSize = LOF(fh) + 1 - currFilePos

'MsgBox "chunksize: " & myChunkSize

If (myChunkSize > CHUNK_SIZE) Then myChunkSize = CHUNK_SIZE

If (myChunkSize < 0) Then
    Call MsgBox("ERROR, myChunkSize < 0!!! impossible to happen..", vbCritical, "Error")
    get_next_chunk = True
    Exit Function
ElseIf (currFilePos + myChunkSize - 1 = ftpFileSize) Then
    get_next_chunk = True
Else
    get_next_chunk = False
End If

currFilePos = currFilePos + myChunkSize

bb_len = myChunkSize

'MsgBox "b4 redim"

ReDim bb(myChunkSize - 1) As Byte

'MsgBox "after redim"

Get #fh, , bb

Close #fh

'Call update_buf(MODE.STATE, myChunkSize & " bytes sent")
pbs.Value = currFilePos - 1
lbytes.Caption = (currFilePos - 1) & " / " & ftpFileSize & " bytes"


End Function


Public Function handle(ByRef data As String, ByRef length As Long, ByRef buf As String, _
    ByRef con As Control, index As Integer) As Integer

    Dim pos As Long
    Dim list As Variant
    Dim counter As Integer
    
    If (length = -1) Then
        pos = InStr(data, DELIMITER)
        length = Val(Left(data, pos - 1))
        
        If length < 1 Then
            handle = RES.NONE
            Exit Function
        End If
        
        buf = Mid(data, pos + 1)
    Else
        'meaning server hasn't finished sending us everything
        buf = buf & data
    End If
    
    Dim s As String
    
    If (Len(buf) = length Or Len(buf) > length) Then
        'myState = CONN
        'output the list of drives
        list = Split(buf, DELIMITER)
        If (index = -1) Then
            con.Clear
            For counter = 0 To UBound(list)
                s = list(counter)
                con.AddItem (s)
                
                'If (max < TextWidth(s & "  ")) Then
                '    max = TextWidth(s & "  ")
                'End If
 
            Next counter
            If (con.ListCount > 0) Then con.ListIndex = 0
            
        Else
 '       MsgBox ("numspaces: " & numSpaces)
         '   For counter = 0 To UBound(list)
            For counter = UBound(list) To 0 Step -1
                s = String(numSpaces, " ") & list(counter)
                con.AddItem s, index
                
               ' If (max < TextWidth(s & "  ")) Then
               '     max = TextWidth(s & "  ")
               ' End If
         
            Next counter
            
        End If
        
        If Len(buf) = length Then
            handle = RES.SUCCESS
        Else
            handle = RES.OVER
        End If
    'ElseIf (Len(buf) > length) Then
       
    '    handle = RES.OVER
    Else
  
        update_log ("Len(buf): " & Len(buf) & " length: " & length)
        handle = RES.UNDER
    End If
    
    'update_scrollbars

End Function



Public Sub request_file(path As String)



    socket_send (PI & path)
    

End Sub
 
Public Sub update_log(s As String)

    tbl.Text = tbl.Text + s + vbCrLf
    tbl.SelStart = Len(tbl.Text)

End Sub
 


Private Sub reset()

    ftp_init

    'tbip.Text = ws.LocalIP
    'tbp.Text = "8888"
    tbl.Text = ""
    
    myState = DORMANT
    'want_files_after_folders = False
   ' first_request = True

    ReDim expansionList(0) As String
    
    cbsd.Clear
    lbf.Clear
    lbfi.Clear
    
    
    reset_drives
    reset_folders
    reset_files

    
    On Error Resume Next

    Drive1.Drive = "C:"
  '  Dir1.Path = "C:\Documents and Settings\t0916\орн▒\vb tutorial"
    Dir1.path = App.path
    File1.path = Dir1.path
    
    If ERR.Number = 76 Then
        MsgBox "Preloaded path is invalid", vbCritical, "Path not found"
    End If
    
    If ws.STATE <> sckClosed Then ws.Close
    
    update_scrollbars

End Sub




Private Sub cbc_Click()


On Error GoTo Error:

'save ip and port into registry
SaveSetting App.Title, "client", "ip", tbip.Text
SaveSetting App.Title, "client", "port", tbp.Text

ws.RemoteHost = tbip.Text
ws.RemotePort = Val(tbp.Text)
ws.Connect
myState = ATTEMPT_CONNECT
'update_log ("Connecting to " & ws.RemoteHost & " at port " & ws.RemotePort & "...")
ls.Caption = "Connecting to " & ws.RemoteHost & " at port " & ws.RemotePort & "..."
Exit Sub

Error:
    'update_log "Error connecting: " & ERR.Description
    ls.Caption = "Error connecting: " & ERR.Description
    ws.Close
    myState = DORMANT
    
End Sub





Private Sub cbf_Click()
File1.Refresh
End Sub

'retrieve file or dir from server
Private Sub cbget_Click()

'exit if my state is not connected
If myState <> CONN Then
    update_log "My state is not CONN, so cannot download files"
    Exit Sub
End If

'exit if no dir or no file is chosen
If lbf.ListIndex = -1 Then
    MsgBox "You need to choose the dir in the server.", vbCritical, "Error"
    Exit Sub
End If


'folder path to download
ftpFolderPath = Trim(lbf.list(lbf.ListIndex))

If Right(ftpFolderPath, 1) <> "\" Then ftpFolderPath = ftpFolderPath & "\"

If lbfi.ListIndex <> -1 Then
    defaultFile = lbfi.list(lbfi.ListIndex)
    'note in libfi there are both file size and file name
    defaultFile = Mid(defaultFile, 1, InStrRev(defaultFile, "(") - 2)
    ftpFilePath = ftpFolderPath & defaultFile
Else
    defaultFile = ""
End If

'get the last dir title in lbf (ex: c:\a\b\ yields b)
defaultDir = Mid(ftpFolderPath, InStrRev(ftpFolderPath, _
        "\", Len(ftpFolderPath) - 1) + 1)
'however, if it's c:\ it should yields nothing
If Mid(defaultDir, 2, 1) = ":" Then defaultDir = ""


Again:
pattern = "*.*"
Form2.Show vbModal, Me
If tranOption = TRAN_CANCEL Then ftp_init: Exit Sub

If Right(defaultFile, 1) = "\" Then defaultFile = Left(defaultFile, Len(defaultFile) - 1)
If Right(defaultDir, 1) <> "\" Then defaultDir = defaultDir & "\"

'Dim ans As DOWNLOAD_TYPE
'ans = MsgBox("Yes: file download" & vbCrLf & "No: dir download" & vbCrLf & _
'    "Cancel: cancel download", vbYesNoCancel, "Choice")
Dim ans As String
Dim ts As String

If tranOption = TRAN_FILE Then
    If lbfi.ListIndex = -1 Then
        MsgBox "You need to choose file in the server.", vbCritical, "Error"
        ftp_file_init
        Exit Sub
    End If
    
   ' If defaultFile = "" Or Len(defaultFile) < 1 Then
   '     MsgBox "Invalid file name; next time know better.", vbCritical, "Error"
   '     ftp_file_init
   '     Exit Sub
   ' End If
    
    'file name on the server to download
   ' Dim fn As String
   ' fn = lbfi.list(lbfi.ListIndex)
    'note in libfi there are both file size and file name
  '  fn = Mid(fn, 1, InStrRev(fn, "(") - 2)
    
    
    
   
    ts = Trim(Dir1.list(Dir1.ListIndex))
    If Right(ts, 1) <> "\" Then ts = ts & "\"
    
    If defaultFile = "" Or Len(defaultFile) < 1 Then
        MsgBox "Invalid file name", vbCritical, "Error"
        ftp_file_init
        Exit Sub
    End If
    
     myFilePath = ts & defaultFile
     
    ' Dim fs As New FileSystemObject
     If fs.FileExists(myFilePath) Then
         ans = MsgBox(myFilePath & " exists. Overwrite?", vbYesNoCancel, "File Exists")
         
         If ans = vbYes Then
             fs.DeleteFile (myFilePath)
         ElseIf ans = vbNo Then
             GoTo Again
         Else
             update_log "File download canceled by user."
             ftp_file_init
             Exit Sub
         End If
     End If
         
     myState = REQ_FILE
     socket_send (PL & ftpFilePath)
     update_log ("Request " & ftpFilePath & " from server...")
        
  '  End If
    
    
    
ElseIf tranOption = TRAN_DIR_R Then
    

    myBasePath = Trim(Dir1.list(Dir1.ListIndex))
    If Right(myBasePath, 1) <> "\" Then myBasePath = myBasePath & "\"
    
    'however, i want their last folder appended to myBasePath. for example
    'myBasePath is C:\doc\ and their last folder is abc\, I want all files under
    'abc\ to store in C:\doc\abc\
   ' myBasePath = myBasePath & Mid(ftpFolderPath, InStrRev(ftpFolderPath, _
    '    "\", Len(ftpFolderPath) - 1) + 1)
    If defaultDir = "" Or Len(defaultDir) < 1 Then
        MsgBox "Invalid dir name", vbCritical, "Error"
        ftp_dir_init
        Exit Sub
    End If
    
    myBasePath = myBasePath & defaultDir
    
    If fs.FolderExists(myBasePath) Then
        ans = MsgBox(myBasePath & " exists. Continue?", vbYesNo, "Folder Exists")
        If ans = vbNo Then
            update_log "Recursive folder download request canceled."
            ftp_init
            Exit Sub
        End If
    End If
    
    myState = REQ_DIR
    socket_send (PZ & "0" & ftpFolderPath & DELIMITER & pattern)
    update_log ("Requesting the entire " & ftpFolderPath & " from server...")


ElseIf tranOption = TRAN_DIR Then

    myBasePath = Trim(Dir1.list(Dir1.ListIndex))
    If Right(myBasePath, 1) <> "\" Then myBasePath = myBasePath & "\"
    
    'however, i want their last folder appended to myBasePath. for example
    'myBasePath is C:\doc\ and their last folder is abc\, I want all files under
    'abc\ to store in C:\doc\abc\
    'myBasePath = myBasePath & Mid(ftpFolderPath, InStrRev(ftpFolderPath, _
    '    "\", Len(ftpFolderPath) - 1) + 1)
        
    'defaultDir = Mid(ftpFolderPath, InStrRev(ftpFolderPath, _
    '    "\", Len(ftpFolderPath) - 1) + 1)
    
    If defaultDir = "" Or Len(defaultDir) < 1 Then
        MsgBox "Invalid dir name", vbCritical, "Error"
        ftp_dir_init
        Exit Sub
    End If
    
    myBasePath = myBasePath & defaultDir
    
    If fs.FolderExists(myBasePath) Then
        ans = MsgBox(myBasePath & " exists. Continue?", vbYesNo, "Folder Exists")
        If ans = vbNo Then
            update_log "Unrecursive folder download request canceled."
            ftp_init
            Exit Sub
        End If
    End If
    
    myState = REQ_DIR
    socket_send (PZ & "1" & ftpFolderPath & DELIMITER & pattern)
    update_log ("Requesting the entire " & ftpFolderPath & " from server...")

ElseIf tranOption <> TRAN_CANCEL Then
    MsgBox "critical error, tranOption is invalid: " & tranOption
    ftp_init
End If


End Sub



'used to upload files to server
Private Sub cbput_Click()

'update_log "in cbput_click(): 1"

On Error Resume Next
'exit if my state is not connected
If myState <> CONN Then
    update_log "My state is not CONN, so cannot upload files"
    Exit Sub
End If

'u know what? Dir1.ListIndex counts from -1!!! But you still can get it's item
'according to my last experiment, the lowest file is -1, the file above it is -2 and so on
'MsgBox Dir1.ListIndex & " dir1: " & Trim(Dir1.list(Dir1.ListIndex))
'exit if no dir or no file is chosen
'If Dir1.ListIndex = -1 Then
'MsgBox Dir1.ListIndex & " dir1: " & Trim(Dir1.list(Dir1.ListIndex))
'    MsgBox "You need to choose the dir on the left side.", vbCritical, "Error"
'    Exit Sub
'End If

ftp_file_init

'ftpFolderPath = Dir1.path
ftpFolderPath = Dir1.list(Dir1.ListIndex)

If Right(ftpFolderPath, 1) <> "\" Then ftpFolderPath = ftpFolderPath & "\"

If File1.ListIndex <> -1 Then
    defaultFile = File1.FileName
Else
    defaultFile = ""
End If



'update_log "in cbput_click(): 2"


'however, file1.listindex is correct...
'MsgBox "file1.listindex: " & File1.ListIndex & " " & defaultFile

'get the last dir title in dir1 (ex: c:\a\b\ yields b\)
defaultDir = Mid(ftpFolderPath, InStrRev(ftpFolderPath, _
        "\", Len(ftpFolderPath) - 1) + 1)
'dir should end with \
If Right(defaultDir, 1) <> "\" Then defaultDir = defaultDir & "\"
'however, if it's c:\ it should yields nothing
If Mid(defaultDir, 2, 1) = ":" Then defaultDir = ""

Again:
pattern = "*.*"
Form2.Show vbModal, Me
If tranOption = TRAN_CANCEL Then ftp_init: Exit Sub

Dim ans As String, ts As String, path As String
Dim index As Long



'update_log "in cbput_click(): 3"

If tranOption = TRAN_FILE Then
    If File1.ListIndex = -1 Then
        MsgBox "You need to choose a file on your pc.", vbCritical, "Error"
        ftp_init
        Exit Sub
    End If
    
    'file name on the client to upload
    myFilePath = ftpFolderPath & File1.FileName
    
    If defaultFile = "" Or Len(defaultFile) < 1 Then
        MsgBox "Invalid file name", vbCritical, "Error"
        ftp_init
        Exit Sub
    End If
    
    If lbf.ListIndex = -1 Then
        MsgBox "You need to select a targeting dir on the server."
        ftp_init
        Exit Sub
    End If
    
    ts = Trim(lbf.list(lbf.ListIndex))
    If Right(ts, 1) <> "\" Then ts = ts & "\"
    
    ftpFilePath = ts & defaultFile
    ftpFileSize = get_size(myFilePath)
    pbs.max = ftpFileSize
    myState = ATTEMPT_SEND
    speedTimer.Enabled = True
    socket_send (PB & ftpFilePath & DELIMITER & ftpFileSize)
    update_log "Wanting to upload " & myFilePath & " to server..."

ElseIf tranOption = TRAN_DIR_R Or TRAN_DIR Then
'we take a different approach to transfer dir to server
'different from the way server transfers dir to client
'we find all paths we need to transfer, then use a for loop to transfer them
'one by one. at first it seems impossible, but we can use doEvents to achieve it
    If lbf.ListIndex = -1 Then
        MsgBox "You need to select a targeting dir on the server."
        ftp_init
        Exit Sub
    End If
    
    If defaultDir <> "" And Right(defaultDir, 1) <> "\" Then
        defaultDir = defaultDir & "\"
    End If
    
    ftpFolderPath = Trim(lbf.list(lbf.ListIndex))
    If Right(ftpFolderPath, 1) <> "\" Then ftpFolderPath = ftpFolderPath & "\"
    myFolderPath = Dir1.list(Dir1.ListIndex)
    If Right(myFolderPath, 1) <> "\" Then myFolderPath = myFolderPath & "\"
    
    If tranOption = TRAN_DIR_R Then
        store_all_files_r myFolderPath, pattern
    Else
        store_all_files myFolderPath, pattern
    End If
    
'    MsgBox allFiles(1)
'On Error Resume Next

'update_log "in cbput_click(): 4"

   ' MsgBox "ubound: " & UBound(allFiles)
  ' update_log "Begin uploading " & myFolderPath & " to the server..."
    For index = 1 To UBound(allFiles)
       ' update_log "in the array!!! " & index & " " & UBound(allFiles)
Again2:
'NEXT LINE causes error when i restart server
'i know why..last round i stopped ftp cold and that means myFilePath didn't get
'init for next connection
'update_log "allfiles's ubound: " & UBound(allFiles) & " index: " & index
        myFilePath = allFiles(index)
        ftpFileSize = get_size(myFilePath)
        If ftpFileSize < 1 Then
            update_log "Size of " & myFilePath & " is below 1 byte, so discard it."
            index = index + 1
            If index > UBound(allFiles) Then
                Exit Sub
            End If
            update_log "just before goto again2 "
            GoTo Again2
        End If
        ftpFilePath = ftpFolderPath & defaultDir & Mid(myFilePath, Len(myFolderPath) + 1)
        pbs.max = ftpFileSize
        myState = ATTEMPT_SEND
        socket_send (PB & ftpFilePath & DELIMITER & ftpFileSize)
        speedTimer.Enabled = True
        update_log "Wanting to upload " & myFilePath & " to server..."
        
        Do Until myState = CONN
            DoEvents
        Loop
        ftp_file_init
    Next
    update_log "Folder, " & myFolderPath & ", upload complete."
    ftp_init
ElseIf tranOption <> TRAN_CANCEL Then
    MsgBox "critical error, tranOption is invalid: " & tranOption
    ftp_init
End If

End Sub

Private Sub cbr_Click()
Dir1.Refresh
End Sub

Private Sub cbs_Click()

If ws.STATE <> sckClosed Then ws.Close
ls.Caption = "Login stopped."

End Sub

Private Sub cbsd_Click()

'MsgBox ("in cbsd_click")

'send request to get folders for the selected drive

'Dim drive As String
'drive = cbsd.list(cbsd.ListIndex)
'    socket_send(PF & drive)
'    update_log ("Request for folders sent")

'If (first_request) Then MsgBox ("shit"): first_request = False: Exit Sub
If myState <> CONN Then
    update_log "My state is not CONN, so cannot retrieve drives: " & myState
    'MsgBox "My state is not CONN, so cannot retrieve folders"
    Exit Sub
End If

myState = WANT_DIRS
'request_folder = cbsd.list(cbsd.ListIndex)
socket_send (PF & cbsd.list(cbsd.ListIndex))
update_log ("Request for folders sent")
filesRequestPath = cbsd.list(cbsd.ListIndex)
   ' want_files_after_folders = True
    

End Sub


Private Sub cbt_Click()

'save settings to registry
SaveSetting App.Title, "client", "drive", Drive1.Drive
SaveSetting App.Title, "client", "dir", Dir1.path
SaveSetting App.Title, "client", "file", File1.path

socket_send (PQ)
'u no what? sometimes doEvents is necessary. in this case, reset() will close socket
'for some reason that PQ won't get sent out if u dont include DoEvents next line
DoEvents
reset
before_connect_controls
'Form_Load

End Sub




Private Sub Dir1_Change()

File1.path = Dir1.path

End Sub



Private Sub Dir1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then PopupMenu mfolder

End Sub



Private Sub Drive1_Change()


On Error Resume Next



            
            
'Dir1.Path = "c:\"
'MsgBox ("drive: " & Dir1.Path)
Dir1.path = Drive1.Drive





'If Left(Drive1.Drive, 1) = "c" Then
'Dir1.Path = "C:\Documents and Settings\t0916\орн▒\vb tutorial"
'End If

End Sub







Private Sub File1_DblClick()

open_file_or_dir File1.path & "\" & File1.FileName

End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = vbRightButton Then PopupMenu mfile

End Sub


Sub load_settings()

Dim ts As String
ts = GetSetting(App.Title, "client", "ip")
If ts = "" Then
    tbip.Text = ws.LocalIP
Else
    tbip.Text = ts
End If

ts = GetSetting(App.Title, "client", "port")
If ts = "" Then
    tbp.Text = "8888"
Else
    tbp.Text = ts
End If

On Error GoTo Error
ts = GetSetting(App.Title, "client", "drive")
Drive1.Drive = IIf(ts = "", "C:", ts)

ts = GetSetting(App.Title, "client", "dir")
Dir1.path = IIf(ts = "", "C:", ts)

ts = GetSetting(App.Title, "client", "file")
File1.path = IIf(ts = "", "C:", ts)
Exit Sub

Error:
Drive1.Drive = "C:"
    
End Sub



Private Sub Form_Load()
   ' Dim fs As New FileSystemObject
   ' Dim dr As Drive
    
   ' For Each dr In fs.Drives
        'List1.AddItem dr.Path & " = " & GetDriveName(dr.DriveType)
        
  '  Next
    initialFormWidth = Form1.Width
    initialFormHeight = Form1.Height
    initialtblWidth = tbl.Width
    initialtblHeight = tbl.Height
    reset
    
    'Dim i As Integer
    'For i = 0 To Drive1.ListCount - 1
    'MsgBox (Drive1.list(i))
    'Next
    
   ' Dim drives As drives
   ' Dim fs As New FileSystemObject
   ' Set drives = fs.drives
   ' For i = 1 To drives.Count
   '     MsgBox (drives.Item(i))
   ' Next
    before_connect_controls
    
    load_settings
    
    fileFoundOption = UNDECIDED
    
End Sub

Private Sub Form_Resize()

Line1.X1 = 0
Line1.X2 = Form1.ScaleWidth
Line1.Y1 = 0
Line1.Y2 = 0
Line2.X1 = Line1.X1
Line2.X2 = Line1.X2
Line2.Y1 = Line1.Y1 + 25
Line2.Y2 = Line1.Y2 + 25
  
If Form1.Width > initialFormWidth Then
    tbl.Width = initialtblWidth + Form1.Width - initialFormWidth
End If
  
If Form1.Height > initialFormHeight Then
    tbl.Height = initialtblHeight + Form1.Height - initialFormHeight
End If

End Sub

Private Sub lisd_DblClick()

'send request to get files

End Sub

'Private Sub lbf_Click()
Private Sub update_scrollbars()
'MsgBox ("lbf_click")

Dim counter As Integer
Dim s As String
Dim max As Integer
max = 0

'update lbf's horizontal scrollbar
For counter = 0 To lbf.ListCount - 1
    s = lbf.list(counter)
    If (max < TextWidth(s & "  ")) Then max = TextWidth(s & "  ")
Next


If ScaleMode = vbTwips Then
    max = max / Screen.TwipsPerPixelX  ' if twips change to pixels
    SendMessageByNum lbf.hwnd, LB_SETHORIZONTALEXTENT, max, 0
End If


max = 0

'update lbfi's horizontal scrollbar
For counter = 0 To lbfi.ListCount - 1
    s = lbfi.list(counter)
    If (max < TextWidth(s & "  ")) Then max = TextWidth(s & "  ")
Next


If ScaleMode = vbTwips Then
    max = max / Screen.TwipsPerPixelX  ' if twips change to pixels
    SendMessageByNum lbfi.hwnd, LB_SETHORIZONTALEXTENT, max, 0
End If

End Sub



Private Sub Form_Terminate()
'MsgBox "terminate"
End Sub

Private Sub Form_Unload(Cancel As Integer)

'save settings to registry
SaveSetting App.Title, "client", "drive", Drive1.Drive
SaveSetting App.Title, "client", "dir", Dir1.path
SaveSetting App.Title, "client", "file", File1.path


socket_send (PQ)
'u no what? sometimes doEvents is necessary. in this case, reset() will close socket
'for some reason that PQ won't get sent out if u dont include DoEvents next line
DoEvents
reset
before_connect_controls

'MsgBox "unload"
End Sub

Private Sub lbf_DblClick()

If myState <> CONN Then
    update_log "My state is not CONN, so cannot retrieve folders"
    'MsgBox "My state is not CONN, so cannot retrieve folders"
    Exit Sub
End If


On Error Resume Next

'if lbf.listindex is 0, that means the fist item is selected
'but the first item is always the drive dir, but we need to get files
If lbf.ListIndex = 0 Then
    filesRequestPath = lbf.list(0)
    myState = WANT_FILES
    socket_send (PI & filesRequestPath)
    Exit Sub
End If

    Dim path As String
    path = lbf.list(lbf.ListIndex)
    
    'we need to know how many spaces in the front of the dir to add more spaces
    'in dirs of deeper level
    numSpaces = Len(path) - Len(Trim(path)) + DIR_OFF
    
    If (in_expansion_list(Trim(path))) Then
  '  MsgBox ("in expansion")
    
        'take path out of expansion list
        remove_from_expansion_list (Trim(path))
        
        'remove all its sub dirs and exit sub
        Dim tn As Integer
        Dim path2 As String
        tn = Len(lbf.list(lbf.ListIndex + 1)) - Len(Trim(lbf.list(lbf.ListIndex + 1)))
        Do While tn >= numSpaces
   ' MsgBox ("tn = numspaces")
            path2 = lbf.list(lbf.ListIndex + 1)
            If (in_expansion_list(Trim(path2))) Then
                remove_from_expansion_list (Trim(path2))
            End If
            
            lbf.RemoveItem (lbf.ListIndex + 1)
            
            tn = Len(lbf.list(lbf.ListIndex + 1)) - Len(Trim(lbf.list(lbf.ListIndex + 1)))
        Loop
        
        
        'even tho we dont get folders, we still need to get files
        filesRequestPath = Trim(path)
        myState = WANT_FILES
        socket_send (PI & filesRequestPath)
        
       ' update_scrollbars
        'Exit Sub
        
    Else
        'put it in the expansion list
        add_to_expansion_list (Trim(path))
        filesRequestPath = Trim(path)
        'MsgBox ("lbf_DblClick: filesRequestPath: " & filesRequestPath)
        myState = WANT_DIRS_A
        socket_send (PF & filesRequestPath)
    End If
   
    
    
    
    
    
    'update_scrollbars
    

End Sub



Private Sub mabout_Click()

MsgBox "This application is the FTP client that communicates with tsFTPd" & _
   vbCrLf & "to enable file transferring between two computers." & vbCrLf & vbCrLf & _
   vbCrLf & "rewritten by Thomas Schilb", vbInformation, "About"

End Sub

Private Sub mdelete_Click()

If File1.ListIndex = -1 Then Exit Sub

Dim fn As String
fn = File1.path & "\" & File1.FileName
Dim ans As String
ans = MsgBox("You sure you wanna delete it?", vbYesNo, "You sure?")
If ans = vbYes Then
    Dim f As File
    Set f = fs.GetFile(fn)
    f.Delete
End If
File1.Refresh

End Sub

Private Sub mdelete2_Click()

Dim ans As String
ans = MsgBox("You sure you wanna delete " & Dir1.list(Dir1.ListIndex) & _
    "?", vbYesNo, "You sure?")

If ans = vbNo Then Exit Sub

fs.DeleteFolder (Dir1.list(Dir1.ListIndex))
Dir1.Refresh
File1.Refresh
Exit Sub
Error:
    MsgBox ERR.Description, vbCritical, "Error"

End Sub

Private Sub mexit_Click()

MsgBox "Thank you for using Wen FTP Client. Please write to " & vbCrLf & _
       "wentaihao@yahoo.com if you have any comments or questions.", vbInformation, "Thank you"
Unload Me

End Sub

Private Sub mexport_Click()



Dim fn As String

On Error GoTo Cancel
    cd.CancelError = True
    cd.DialogTitle = "Save"
    cd.Filter = "All Files (*.*)|*.*"
    cd.InitDir = App.path
    cd.ShowSave 'shows the save dialog box
    fn = cd.FileName

    If (fn = "" Or Len(fn) < 1) Then
        Call MsgBox("Invalid file name.", vbCritical, "Error")
        Exit Sub
    End If

    Dim ans As Integer
    If (fs.FileExists(fn)) Then
        ans = MsgBox(fn & " exists. Overwrite? (No means cancel)", _
            vbYesNo, "File Exists")
            
        If (ans = vbYes) Then
            fs.DeleteFile (fn)
        ElseIf (ans = vbNo) Then
            Exit Sub
        End If
    End If
    
    
    Dim data As String
    'append date
    data = Date & vbCrLf & vbCrLf & tbl.Text
    Dim fh As Long
    fh = FreeFile
    Open fn For Binary As #fh
    Put fh, , data
    Close #fh
    
    update_log "Log exported to " & fn
   ' obStatus.Value = True
    Exit Sub

Cancel:

End Sub

Private Sub mhelp_Click()

MsgBox "First you click on 'Log in' to log into Wen FTP Server, then you " & vbCrLf & _
       "enter login information. If you don't have an account, You need " & vbCrLf & _
       "to contact the server and tell him to add you and your privilege " & vbCrLf & _
       "level. After you log in, depending on your privilege level, you" & vbCrLf & _
       "may download and upload a file or an entire folder from and to" & vbCrLf & _
       "the server. You may also view the entire files system on the " & vbCrLf & _
       "server given the privilege to do so." & vbCrLf & vbCrLf & _
       "Note: " & vbCrLf & _
       "1. This application may not work with languages other than" & vbCrLf & _
       "English, so make sure your files are named in English." & vbCrLf & _
       vbCrLf & vbCrLf & "written by Michael Wen.", vbInformation, "Help"

End Sub

Private Sub mopen_Click()

open_file_or_dir File1.path & "\" & File1.FileName

End Sub

Private Sub mopen2_Click()

MsgBox "dir: " & Dir1.list(Dir1.ListIndex) & " index: " & Dir1.ListIndex
'Dir1.path doesn't give u the path highlighted
'it gives u the path "opened", u can see the icon next to it opened
open_file_or_dir Dir1.list(Dir1.ListIndex)

End Sub

Private Sub mproperties_Click()

If File1.ListIndex = -1 Then Exit Sub

fileForProperty = File1.path & "\" & File1.FileName
'MsgBox "fileForProperty: " & fileForProperty
propertyType = WANT_FILE_PROPERTY
Form4.Show vbModal, Me
File1.Refresh

End Sub

Private Sub mproperties2_Click()

fileForProperty = Dir1.list(Dir1.ListIndex)
'MsgBox "fileForProperty: " & fileForProperty
propertyType = WANT_FOLDER_PROPERTY
Form4.Show vbModal, Me
File1.Refresh



End Sub

Private Sub mrename_Click()

If File1.ListIndex = -1 Then Exit Sub

Dim fn As String
fn = File1.path & "\" & File1.FileName
Dim ans As String
ans = InputBox("Enter new name:", "New Name", fn)
Dim f As File
Set f = fs.GetFile(fn)
On Error GoTo Error
f.Name = ans
File1.Refresh
Exit Sub
Error:
    MsgBox ERR.Description, vbCritical, "Error"

End Sub

Private Sub mrename2_Click()

Dim dir_path As String
dir_path = Dir1.list(Dir1.ListIndex)
Dim ans As String
ans = InputBox("Enter new name:", "New Name", dir_path)

Dim fo As Folder
Set fo = fs.GetFolder(dir_path)
On Error GoTo Error
fo.Name = ans
'fs.MoveFolder Dir1.path, Dir1.path & "2"

'Dir1.Refresh
'File1.Refresh
Exit Sub
Error:
    MsgBox ERR.Description, vbCritical, "Error"

End Sub

Private Sub speedTimer_Timer()

If (counter = 0) Then
    lspeed = "Transferring at 0 KB/s"
Else
    lspeed = "Transferring at " & CLng((currFilePos - 1) / counter / 1024) & " KB/s"
End If
counter = counter + 1
    
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)

'
On Error GoTo ERR

    Dim s As String
    Dim finished As Boolean
    Dim pos As Integer
    Dim counter As Integer
    Dim list As Variant
    Dim result As Integer
    Dim b As Byte
    ReDim bb(bytesTotal - 1) As Byte
    Dim fp As String
    Dim PF As String
    Dim fd As String
    Dim fh As Long
    Dim ans As String
'    Dim fs As New FileSystemObject
    
    Call ws.PeekData(s)
    Call ws.GetData(bb)
    
   ' MsgBox "s: " & s

    
    Select Case myState
        Case ATTEMPT_CONNECT
            If (Left(s, 1) = PP) Then
                ls.Caption = "Connected to server."
                
                load_settings
                
                'show login window
                Form3.Show vbModal, Me
                 
                If loginCancel Then
                'user clicks on cancel, so we tell server to break this connection
                'and listen again
                    socket_send (PQ)
                    ws.Close
                    ls.Caption = "Connection terminated."
                    myState = DORMANT
                Else
                    socket_send (PG & loginName & DELIMITER & loginPassword)
                    ls.Caption = "Login info sent."
                End If
                
            ElseIf (Left(s, 1) = PY) Then
            
                update_log ("Logged into server.")
                after_connect_controls
                
                'get drives so that user can choose a drive
                myState = WANT_DRIVES
                socket_send (PD)
                update_log ("Request for the drive list sent")
                
            ElseIf (Left(s, 1) = PX) Then
                ls.Caption = "Connection closed by server: " & Mid(s, 2)
                ws.Close
                myState = DORMANT
            End If
        
        Case WANT_DRIVES
            If (Left(s, 1) = PX) Then
                update_log "Error from server: " & Mid(s, 2)
                myState = CONN
                Exit Sub
            End If
            
                result = handle(s, drivesLen, drivesBuf, cbsd, -1)
                
                Select Case result
                    Case RES.NONE
                        update_log ("Server's drive list retrieved, but no items in the list.")
            'initially we need to add more code here to get folders from server
            'however it's not needed cuz after update driver list
            'driver1_click will be invoked automatically although logically
            'it shouldn't happen until user physically clicks the combo box
            'anyway, we want drives only when we connect to the server
            'after that we play with folders and files
                        
                      '  myState = WANT_DIRS
                      '  filesRequestPath = cbsd.list(cbsd.ListIndex)
                      '  socket_send(PF & filesRequestPath)
                      '  update_log ("Request for folders sent")
                        reset_drives
                        
                        cbsd.Clear
                        
                    Case RES.SUCCESS
                        update_log ("Server's drive list retrieved.")
                        myState = CONN
                      '  myState = WANT_DIRS
                      '  filesRequestPath = cbsd.list(cbsd.ListIndex)
                      '  socket_send(PF & filesRequestPath)
                      '  update_log ("Request for folders sent")
                        reset_drives
           
                    Case RES.OVER
                        update_log ("We got more stuff than claimed by server." & _
                        vbCrLf & "Len(drivesBuf): " & Len(drivesBuf) & vbCrLf & _
                        "driveLen: " & drivesLen)
                        
                        reset_files
                        myState = CONN
           
                    Case RES.UNDER
                        update_log ("Waiting for more stuff in the drive list...")
                        
                End Select
        'used when user double clicks folder list to get sub folders
        'as u can see sub folders are added to the original folder list, unlike
        'how WANT_DIRS works
        Case WANT_DIRS_A
            If (Left(s, 1) = PX) Then
                update_log "Error from server: " & Mid(s, 2)
                myState = CONN
                Exit Sub
            End If
        
            result = handle(s, foldersLen, foldersBuf, lbf, lbf.ListIndex + 1)
            
            Select Case result
                 Case RES.NONE
                    update_log ("Server's folder list retrieved, but no items in the list.")
                    reset_folders
                    
                    'lbf.Clear
                 '   MsgBox ("should be true: " & in_expansion_list(filesRequestPath))
                        remove_from_expansion_list (filesRequestPath)
                    
                        myState = WANT_FILES
                        'want_files_after_folders = False
                        socket_send (PI & filesRequestPath)
                   
                    
                 Case RES.SUCCESS
                    update_log (filesRequestPath & " Server's folder list retrieved.")
                    reset_folders
                   
                   
                        myState = WANT_FILES
                      '  want_files_after_folders = False
                        socket_send (PI & filesRequestPath)
        
                 Case RES.OVER
                 'the bane of Chinese names...
                     update_log ("We got more stuff than claimed by server." & _
                     vbCrLf & "Len(foldersBuf): " & Len(foldersBuf) & vbCrLf & _
                     "foldersLen: " & foldersLen & vbCrLf & "s: " & s)
                     
                     myState = WANT_FILES
                     socket_send (PI & filesRequestPath)
                     'reset_files
                    'myState = CONN
        
                 Case RES.UNDER
                     update_log ("Waiting for more stuff in the folder list...")
                    
             End Select
             update_scrollbars
             DoEvents
        'when user selects a drive the folders are updated
        'no such concept as "sub folder" here like in WANT_DIRS_A
        Case WANT_DIRS
            If (Left(s, 1) = PX) Then
                update_log "Error from server: " & Mid(s, 2)
                myState = CONN
                Exit Sub
            End If
            
            result = handle(s, foldersLen, foldersBuf, lbf, -1)
            
            Select Case result
                 Case RES.NONE
                    update_log ("Server's folder list retrieved, but no items in the list.")
                    reset_folders
                    
                    lbf.Clear
                    lbf.AddItem cbsd.list(cbsd.ListIndex), 0
                    lbf.ListIndex = 0
                    
                        myState = WANT_FILES
                        'want_files_after_folders = False
                        socket_send (PI & filesRequestPath)
                   
                    
                 Case RES.SUCCESS
                    update_log (filesRequestPath & " Server's folder list retrieved.")
                    reset_folders
                   
                   'we need root folder, or drive folder in the first position
                    lbf.AddItem cbsd.list(cbsd.ListIndex), 0
                    lbf.ListIndex = 0
                        
                        myState = WANT_FILES
                      '  want_files_after_folders = False
                        socket_send (PI & filesRequestPath)
        
                 Case RES.OVER
                     update_log ("We got more stuff than claimed by server." & _
                     vbCrLf & "Len(foldersBuf): " & Len(foldersBuf) & vbCrLf & _
                     "foldersLen: " & foldersLen & vbCrLf & "s: " & s)
                     lbf.AddItem cbsd.list(cbsd.ListIndex), 0
                    lbf.ListIndex = 0
                     myState = WANT_FILES
                    socket_send (PI & filesRequestPath)
                     'reset_files
                    'myState = CONN
        
                 Case RES.UNDER
                     update_log ("Waiting for more stuff in the folder list...")
                     
             End Select
             update_scrollbars
             DoEvents
             
        Case WANT_FILES
            If (Left(s, 1) = PX) Then
                update_log "Error from server: " & Mid(s, 2)
                myState = CONN
                Exit Sub
            End If
            
            result = handle(s, filesLen, filesBuf, lbfi, -1)
            Select Case result
                 Case RES.NONE
                    update_log ("Server's file list retrieved, but no items in the list.")
                    lbfi.Clear
                    reset_files
                    myState = CONN
                    
                 Case RES.SUCCESS
                    update_log (filesRequestPath & " Server's file list retrieved.")
                    reset_files
                    myState = CONN
        
                 Case RES.OVER
                     update_log ("We got more stuff than claimed by server." & _
                     vbCrLf & "Len(filesBuf): " & Len(filesBuf) & vbCrLf & _
                     "filesLen: " & filesLen & vbCrLf & "s: " & s)
                     
                     reset_files
                    myState = CONN
        
                 Case RES.UNDER
                     update_log ("Waiting for more stuff in the file list...")
                    
             End Select
             update_scrollbars
             'MsgBox "b4 files doevents"
             DoEvents
             
        Case REQ_FILE
            
            If Left(s, 1) = PB Then
                'the requested file exists and ready to send to me
                'i need to get size
                ftpFileSize = Val(Mid(s, 2))
                
                pbs.max = ftpFileSize
                
                'we already deal with if-file-exists problem when << is clicked
                'go take a look at that function
                
                
                
                'make sure myFilePath exists, if it doesn't create it
                'fs.GetParentFolderName returns the folder path without ending \
                PF = fs.GetParentFolderName(myFilePath)
                If (Mid(PF, Len(PF)) <> "\") Then PF = PF & "\"
                If (Not fs.FolderExists(PF)) Then
                    pos = InStr(PF, "\")
                    Do While pos <> 0
                        pos = InStr(pos + 1, PF, "\")
                        If (pos = 0) Then Exit Do
                        fd = Mid(PF, 1, pos)
                        If (Not fs.FolderExists(fd)) Then
                            fs.CreateFolder (fd)
                        End If
                    Loop
                End If
                
                update_log "Begin receiving " & myFilePath & " (" & ftpFileSize & _
                    " bytes)..."
                myState = RECV_FILE
                speedTimer.Enabled = True
                socket_send (PA & 0)
                
            ElseIf Left(s, 1) = PX Then
                update_log "Error from server: " & Mid(s, 2)
                ftp_file_init
                myState = CONN
                
                
           ' ElseIf Left(s, 1) = PE Then
                'the requested file doesn't exist
                'MsgBox "Server doesn't have " & ftpFilePath, vbCritical, "Error"
             '   update_log "Server doesn't have " & ftpFilePath
             '   ftp_file_init
              '  myState = CONN
                
            Else
                MsgBox "My state is REQ_FILE but I got: " & s, vbCritical, "Error"
                ftp_file_init
                myState = CONN
                
            End If
            
            
        Case REQ_DIR
            If Left(s, 1) = PB Then
                'the requested folder exists and ready to send to me
                'i need to get file path and its size
                pos = InStr(1, s, DELIMITER)
                fp = Mid(s, 2, pos - 2)
                ftpFileSize = Val(Mid(s, pos + 1))
                
                pbs.max = ftpFileSize
                
                'remember the fp is absolute file path, but i only want the
                'part after base folder, which is stored in ftpFilePath
                'so i use mid() to get the part after ftpFolderPath's length
                'then i attached it to myBasePath
                fp = Mid(fp, Len(ftpFolderPath) + 1)
                myFilePath = myBasePath + fp
                
               ' MsgBox "fileFoundOption 1: " & fileFoundOption
                
                'if file exists, ask user if he wants it or not
                'as u can see we cannot deal with this problem when user clicks <<
                'like file cuz we are ftping folder this time and we don't know
                'the names of files in the folder when user clicks <<
                'we don't know myFilePath until now...so let's deal with it now
                If fs.FileExists(myFilePath) Then
                    If fileFoundOption = UNDECIDED Or _
                        fileFoundOption = DO_NOT_OVERWRITE Or _
                        fileFoundOption = OVERWRITE Or _
                        fileFoundOption = DISCARD Then
                        Form5.Show vbModal
                    End If
                    
                    If fileFoundOption = OVERWRITE Or _
                        fileFoundOption = OVERWRITE_ALL Then
                        fs.DeleteFile (myFilePath)
                    ElseIf fileFoundOption = DO_NOT_OVERWRITE Then
                        'newFilePath must be a valid path, as checked by is_filename_valid
                    MsgBox "new path: " & newFilePath
                        myFilePath = newFilePath
                    Else
                        'user doesn't want this file
                        update_log "I don't want " & myFilePath & "."
                        ftp_file_init
                        socket_send (PA & PA)
                        Exit Sub
                    End If
                    
                End If
                
               ' MsgBox "fileFoundOption 2: " & fileFoundOption
                
                
                'fs.GetParentFolderName returns the folder path without ending \
                PF = fs.GetParentFolderName(myFilePath)
                'i need to create the dir if it doesn't exist, for example, if user
                'wants C:\a\b\c\d\e but I only have C:\a, I need to create C:\a\b
                'then C:\a\b\c, then C:\a\b\c\d, then C:\a\b\c\d\e
                If (Mid(PF, Len(PF)) <> "\") Then PF = PF & "\"
                If (Not fs.FolderExists(PF)) Then
                    pos = InStr(PF, "\")
                    Do While pos <> 0
                        pos = InStr(pos + 1, PF, "\")
                        If (pos = 0) Then Exit Do
                        fd = Mid(PF, 1, pos)
                        If (Not fs.FolderExists(fd)) Then
                            fs.CreateFolder (fd)
                        End If
                    Loop
                End If
                
                update_log "Begin receiving " & myFilePath & " (" & ftpFileSize & _
                    " bytes)..."
                myState = RECV_DIR
                speedTimer.Enabled = True
                socket_send (PA & 0)
                
            ElseIf Left(s, 1) = PX Then
                update_log "Error msg from server: " & Mid(s, 2)
                ftp_file_init
                ftp_dir_init
                myState = CONN
                
                
            ElseIf Left(s, 1) = PE Then
                'the requested file doesn't exist
                'MsgBox "Server doesn't have " & ftpFilePath, vbCritical, "Error"
                update_log "Folder download complete"
                ftp_file_init
                ftp_dir_init
                fileFoundOption = UNDECIDED
                Dir1.Refresh
                File1.Refresh
                myState = CONN
                
            Else
                MsgBox "My state is REQ_FILE but I got: " & s, vbCritical, "Error"
                ftp_file_init
                ftp_dir_init
                myState = CONN
                
            End If
            
            
        Case RECV_FILE
        

        
            finished = recv_chunk(bb, bytesTotal)
            socket_send (PA & bytesTotal)
            If (finished) Then
                myState = CONN
                ftp_file_init
                update_log "File download complete."
                File1.Refresh
                fileFoundOption = UNDECIDED
            End If
            
         
        Case RECV_DIR
        

        
            finished = recv_chunk(bb, bytesTotal)
             socket_send (PA & bytesTotal)
             
            If (finished) Then
                myState = REQ_DIR
                ftp_file_init
                update_log myFilePath & " download complete."
            Else
            
               

            End If
            
        Case ATTEMPT_SEND
        
        ' MsgBox "got PA from server: " & s
        
            If Left(s, 1) = PA Then
            
       ' MsgBox "1"
            
                myState = SEND_FILE
                'ftp_file_init
                
      '  MsgBox "2"
                'send first chunk
                ReDim bb(0) As Byte
                Dim bb_len As Long
                
  '      MsgBox "3"
                finished = get_next_chunk(bb, bb_len)
                
   '     MsgBox "4: " & bb_len
                socket_send (bb)
                
   '     MsgBox "5"
                If (finished) Then
                    update_log "Last packet sent"
                End If
            ElseIf (Left(s, 1) = PX) Then
                update_log "Error from server: " & Mid(s, 2)
                myState = CONN
            Else
                update_log "My state is ATTEMPT_SEND but I got: " & s
            End If
        
        Case SEND_FILE
        
       ' MsgBox "got PA from server1: " & s
        
            If Left(s, 1) = PA Then
            
          '  MsgBox "got PA from server2: " & s
            
                s = Mid(s, 2)
                Dim ss() As String
                ss() = Split(s, PA) 'u dont split it with DELIMITER...dude...wake up!!
                For counter = LBound(ss) To UBound(ss)
                    theirChunkSize = theirChunkSize + Val(ss(counter))
                Next
                
         '   MsgBox "theirChunkSize: " & theirChunkSize
                
                If (myChunkSize = theirChunkSize) Then
                
                    If (currFilePos - 1 = ftpFileSize) Then
                        ftp_file_init
                        update_log "File upload to server complete."
                        
                        If myState = SEND_FILE Then
                            
                            myState = CONN
                        
                        ElseIf myState = SEND_DIR Then
Again:
                            If filePointer > UBound(allFiles) Then
                                ftp_dir_init
                                myState = CONN
                                'send PE
                                socket_send (PE)
                                update_log "Folder upload to server complete."
                                Exit Sub
                            End If
                        
                            'send next file fp points to
                            myFilePath = allFiles(filePointer)
                            filePointer = filePointer + 1
                        
                            If (fs.FileExists(myFilePath)) Then
                          'send "B<fileName excluding ftpFolderPath>DELIMITER<file size>"
                    
                                fh = FreeFile
                                Open myFilePath For Binary As #fh
                                ftpFileSize = LOF(fh)
                                Close #fh
                                
                                If ftpFileSize <= 0 Then
                                   ' socket_send(PX & "File size of " & ftpFilePath & " is 0 byte.")
                                    update_log "File size of " & myFilePath & " is 0 byte; sending next one in line"
                                    GoTo Again
                                Else
                                    pbs.max = ftpFileSize
                                    socket_send (PB & myFilePath & DELIMITER & ftpFileSize)
                                    update_log "Client wants " & myFilePath & " (" & ftpFileSize & " bytes), " & _
                                        "PB sent"
                                    speedTimer.Enabled = True
                                
                                End If
                            Else
                                update_log myFilePath & " doesn't exist; sending next one in line"
                                GoTo Again
                            End If
                    
                        End If
                    
                    
                        Exit Sub
                    
                    End If
                
                
                    myChunkSize = 0
                    theirChunkSize = 0
                    
     
                    ReDim bb(0) As Byte
                   ' Dim bb_len As Long
                    finished = get_next_chunk(bb, bb_len)
    
                    socket_send (bb)
                    
                    If (finished) Then
                        update_log "Last packet sent"
                    End If
            
               ElseIf (myChunkSize > theirChunkSize) Then
                   update_log "My chunk size: " & myChunkSize & " and their chunk " _
                       & "size: " & theirChunkSize & ". Waiting for more Ack..."
               
               Else
                   update_log "My chunk size: " & myChunkSize & " and their chunk " _
                       & "size: " & theirChunkSize & ". Forced to terminate FTP because " & _
                       "server claimed to " & vbCrLf & "have received more than we sent."
                   
                   myState = CONN
                   ftp_file_init
                   s = "You shouldn't have received more than we sent you."
                   socket_send (PX & s)
    
               End If
            Else
                update_log "My state is SEND_FILE but I didn't get PA: " & s
            End If
            
    End Select
        
   
    
    Exit Sub
ERR:
    update_log ("Error: " & ERR.Description)
   

End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

update_log ("Error: " & Description)
reset
    
End Sub
