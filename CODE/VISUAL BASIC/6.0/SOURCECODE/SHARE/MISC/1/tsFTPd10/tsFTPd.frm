VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "tsFTPd"
   ClientHeight    =   6240
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5280
   Icon            =   "tsFTPd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd 
      Left            =   3840
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cbc 
      Caption         =   "&Clear Log"
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
      Left            =   960
      TabIndex        =   15
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Timer logoTimer 
      Interval        =   100
      Left            =   2760
      Top             =   1560
   End
   Begin VB.Timer speedTimer 
      Left            =   3240
      Top             =   1560
   End
   Begin MSComctlLib.ProgressBar pbs 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4440
      TabIndex        =   10
      Top             =   2040
      Width           =   735
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   2040
      Width           =   735
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
      Height          =   3015
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Text            =   "tsFTPd.frx":058A
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox tbmp 
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
      Left            =   960
      TabIndex        =   5
      Text            =   "tbmp"
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox tbmi 
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
      Left            =   960
      TabIndex        =   3
      Text            =   "tbmi"
      Top             =   840
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cbt 
      Caption         =   "&Terminate"
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
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton cbss 
      Caption         =   "&Start Server"
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
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   5280
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5280
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label logo 
      AutoSize        =   -1  'True
      Caption         =   "Cheers!"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   3840
      TabIndex        =   14
      Top             =   1200
      Width           =   855
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
      TabIndex        =   13
      Top             =   5880
      Width           =   2655
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
      Left            =   2880
      TabIndex        =   11
      Top             =   5880
      Width           =   2295
   End
   Begin VB.Label ll 
      Caption         =   "Log:"
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
      TabIndex        =   7
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "My Port:"
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
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "My IP:"
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
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin VB.Menu mfile 
      Caption         =   "&File"
      Begin VB.Menu mexport 
         Caption         =   "E&xport Log"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnone 
         Caption         =   "-"
      End
      Begin VB.Menu mexit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mabout2 
      Caption         =   "&About"
      Begin VB.Menu mabout 
         Caption         =   "A&bout tsFTPd"
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
'note that my DELIMITER must be the same as client's DELIMITER
'wrong: DELIMITER cannot be : cuz it's used in file list to separate file size from name

'be careful with using On error resume next because it ignores errors without
'doing anything. so if ur code should print sth but it doesn't, it probably means
'what it's printing has an error and is ignored


Option Explicit

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


'make sure myFilePath's base dir exists before calling this function
'for ex, if myFilePath is C:\a\b\c\d.txt, C:\a\b\c must exist
Public Function recv_chunk(bb() As Byte, bb_len As Long) As Boolean

'MsgBox "in recv_chunk"

If (currFilePos <= 0) Then currFilePos = 1

'MsgBox "1"

Dim fh As Long
fh = FreeFile
If myFilePath = "" Or Len(myFilePath) < 1 Then
    MsgBox ("Impossible error in recv_chunk: myFilepath is invalid")
    Exit Function
End If

'MsgBox "2"

Open myFilePath For Binary As fh

'MsgBox "3"

Seek fh, LOF(fh) + 1

'MsgBox "left: " & (currFilePos + bb_len - 1) & " right: " & ftpFileSize
    
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
   
If (currFilePos <= 0) Then currFilePos = 1

Dim fh As Long
fh = FreeFile
Open myFilePath For Binary As #fh
Seek fh, currFilePos

myChunkSize = LOF(fh) + 1 - currFilePos

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

ReDim bb(myChunkSize - 1) As Byte
Get #fh, , bb

Close #fh

'Call update_buf(MODE.STATE, myChunkSize & " bytes sent")
pbs.Value = currFilePos - 1
lbytes.Caption = (currFilePos - 1) & " / " & ftpFileSize & " bytes"


End Function
    
Public Sub update_log(s As String)

    tbl.Text = tbl.Text + s + vbCrLf
    tbl.SelStart = Len(tbl.Text)

End Sub

Private Sub reset()

AccountFile = App.path & "/accounts.txt"

ftp_file_init
ftp_dir_init

Drive1.Visible = False
Dir1.Visible = False
File1.Visible = False

    cbss.Enabled = True
    cbt.Enabled = False
    tbmi.Text = ws.LocalIP
    tbmi.Enabled = False
    'tbmp.Text = "8888"
    tbmp.Enabled = True
    'tbl.Text = ""
    lspeed = ""
    lbytes = ""
    
    myState = DORMANT
    
    On Error Resume Next

    Drive1.drive = "C:"
  '  Dir1.Path = "C:\Documents and Settings\t0916\орн▒\vb tutorial"
    Dir1.path = App.path
    File1.path = Dir1.path
    
    If ERR.Number = 76 Then
        MsgBox "Preloaded path is invalid", vbCritical, "Path not found"
    End If
    
    If ws.STATE <> sckClosed Then ws.Close

End Sub




Private Sub cbc_Click()
tbl.Text = ""
End Sub

Private Sub cbss_Click()

On Error GoTo ERR
    
    ws.LocalPort = Val(tbmp.Text)
    ws.Listen
    update_log ("Listening on port " & ws.LocalPort & "...")
    cbss.Enabled = False
    cbt.Enabled = True
    
    SaveSetting App.Title, "server", "port", tbmp.Text
    Exit Sub
    
ERR:
    update_log ("Error: " & ERR.Description)
    ws.Close

End Sub

Private Sub cbt_Click()

    update_log ("Server terminated.")
   
    reset

End Sub




Private Sub Form_Load()
    
initialFormWidth = Form1.Width
initialFormHeight = Form1.Height

    reset
'reset() doesn't clear log
tbl.Text = ""

logo.Caption = ""
logoTimer.Enabled = False
logoTimer.Interval = 10
logoDegree = 0

Dim ts As String
ts = GetSetting(App.Title, "server", "port")
If ts = "" Then
    tbmp.Text = "2121"
Else
    tbmp.Text = ts
End If

'check to see if accounts file exists
If Not fs.FileExists(AccountFile) Then
    MsgBox "Account file " & AccountFile & " does not exist, exiting."
    Unload Me
End If

End Sub

Private Sub Form_Resize()

Line1.X1 = 0
Line1.X2 = Me.ScaleWidth
Line1.Y1 = 0
Line1.Y2 = 0
Line2.X1 = Line1.X1
Line2.X2 = Line1.X2
Line2.Y1 = Line1.Y1 + 25
Line2.Y2 = Line1.Y2 + 25

If Form1.Width > initialFormWidth Then
    tbl.Width = Form1.ScaleWidth - UNIT * 2
    pbs.Width = tbl.Width
End If

If Form1.Height > initialFormHeight Then
    tbl.Height = Form1.ScaleHeight - ll.Top - ll.Height - pbs.Height - _
        lspeed.Height - UNIT * 4
    pbs.Top = tbl.Top + tbl.Height + UNIT
    lspeed.Top = pbs.Top + pbs.Height
    lbytes.Top = lspeed.Top
End If

End Sub

Private Sub logoTimer_Timer()

logoDegree = logoDegree + 1
If logoDegree > 359 Then logoDegree = 0
Dim x As Double, y As Double
x = Math.Cos(logoDegree * PIE / 180) * A
'y = Math.Sqr((1 - (x * x) / (A * A)) * B * B)
y = Math.Sin(logoDegree * PIE / 180) * b
logo.Left = 3840 + x
logo.Top = 1200 + y
End Sub

Private Sub mabout_Click()

MsgBox "This FTPd communicates with tsFTP Client" & _
   vbCrLf & "to enable file transferring between two computers." & vbCrLf & vbCrLf & _
   vbCrLf & "rewritten by Thomas Schilb", vbInformation, "About"

End Sub

Private Sub mexit_Click()

MsgBox "Thank you for using tsFTPd. Please write to " & vbCrLf & _
       "thomas_schilb@outlook.com if you have any comments or questions.", vbInformation, "Thank you"
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

MsgBox "First you click on 'Start Server' to start running the FTPd" & vbCrLf & _
       "server. You have a file named 'accounts.txt in which you " & vbCrLf & _
       "keep client account information. You define its login name," & vbCrLf & _
       "login password, and privilege level." & vbCrLf & vbCrLf & _
       vbCrLf & vbCrLf & "rewritten by Thomas Schilb.", vbInformation, "Help"

End Sub

Private Sub speedTimer_Timer()

If (counter = 0) Then
        lspeed = "Transferring at 0 KB/s"
    Else
        lspeed = "Transferring at " & CLng((currFilePos - 1) / counter / 1024) & " KB/s"
    End If
    counter = counter + 1
    
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)

On Error GoTo ERR

' Accept the incoming connection from the client.
   ws.Close 'we need to close socket before we accept connection
   ws.Accept (requestID)
   myState = CONN
   'tbl.Text = tbl.Text + "Incoming connection accepted." + vbCrLf
   update_log ("Incoming connection accepted: " & ws.RemoteHostIP)

   socket_send (PP)
   
Exit Sub

ERR:
   ' tbl.Text = tbl.Text + "Error: " & Err.Description + vbCrLf
    update_log ("Error: " & ERR.Description)

End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)

On Error Resume Next

'
Dim b As Byte
Dim counter As Integer
ReDim bb(bytesTotal - 1) As Byte
Dim s As String
Dim finished As Boolean
Dim fh As Long
'Dim fs As New FileSystemObject
Dim pattern As String
Dim pos As Long
Dim pfn As String
Dim fd As String
            
Call ws.PeekData(s)
Call ws.GetData(bb)

'MsgBox "s: " & s

On Error Resume Next

'MsgBox ("request from client: " & s)

'if i m receiving a file or dir, i should get that chunk and exit sub
'maybe in the future i should determine actions by my state, not the protocol i received
Select Case myState
    Case RECV_FILE
        'MsgBox " in recv_file: " & bytesTotal
        
        finished = recv_chunk(bb, bytesTotal)
        socket_send (PA & bytesTotal)
        If (finished) Then
            myState = CONN
            ftp_file_init
            update_log "File upload complete."
        End If
        Exit Sub
        
    Case RECV_DIR
    
        finished = recv_chunk(bb, bytesTotal)
         socket_send (PA & bytesTotal)
        If (finished) Then
            myState = REQ_DIR
            ftp_file_init
            update_log myFilePath & " upload complete."
        Else

        End If
        Exit Sub
End Select


Select Case Left(s, 1)

    'get drivers' list from server
    Case PD
        If myState = CONN Then
            If (clientPermission And PERMIT_VIEW_DRIVES) <> 0 Then
    
                Dim driveList As String
                driveList = ""
                
                Drive1.Refresh
                
                If (Drive1.ListCount < 1) Then
                    socket_send ("0" & DELIMITER)
                    update_log ("Drive list (no data) sent to " & clientName)
                    Exit Sub
                End If
                
                For counter = 0 To Drive1.ListCount - 1
                 '  MsgBox (Drive1.List(counter))
                   driveList = driveList & Drive1.List(counter) & DELIMITER
                Next
                'strip the last #
                driveList = Left(driveList, Len(driveList) - 1)
                s = Len(driveList) & DELIMITER & driveList
                socket_send (s)
                update_log ("Drive list sent to " & clientName)
            
            Else
            
                socket_send (PX & "You have no permission to view drives.")
                update_log clientName & " wants to view drives but doesn't have permission."
            
            End If
        
        Else
            update_log "My state is not CONN but I got PD"
        End If
        
    'get dirs from server
    Case PF
        If myState = CONN Then
            If (clientPermission And PERMIT_VIEW_DIRS) <> 0 Then
                'get drive
                Dim drive As String
                drive = Mid(s, 2)
        
                If (Right(drive, 1) <> "\") Then drive = drive & "\"
                
                Dir1.path = drive
                
                Dir1.Refresh
                
                Dim dp As String
                dp = Dir1.path
                If Right(dp, 1) <> "\" Then dp = dp & "\"
        
                'u no what's funny? if i assign a dir to dir1 or file1, they look at it
                'and if it's not available, they will remain the same...
                'why the hell does MS do this to us?
                If (Dir1.ListCount < 1 Or dp <> drive) Then
                    socket_send ("0" & DELIMITER)
                    update_log "Folder list of " & dp & " (no data) sent to " & clientName
                    Exit Sub
                End If
                
                Dim folderList As String
                folderList = ""
                For counter = 0 To Dir1.ListCount - 1
                    folderList = folderList & Dir1.List(counter) & DELIMITER
                Next
                folderList = Left(folderList, Len(folderList) - 1)
                s = Len(folderList) & DELIMITER & folderList
                
                socket_send (s)
                update_log "Folder list of " & dp & " sent to " & clientName
            
            Else
                socket_send (PX & "You have no permission to view folders.")
                update_log clientName & " wants to view folders but doesn't have permission."
            End If
        
        Else
            update_log "My state is not CONN but I got PF"
        End If
    
    'get file list from server
    Case PI
        If myState = CONN Then
        'MsgBox "permit fiels: " & (clientPermission And PERMIT_VIEW_FILES)
            If (clientPermission And PERMIT_VIEW_FILES) <> 0 Then
                'get dir
                Dim dir As String
                dir = Mid(s, 2)
                If (Right(dir, 1) <> "\") Then dir = dir & "\"
                
                File1.path = dir
                'as i said, if dir is not available, file1.path remains the same
                'a perculiar habit of file1.path is that if it's a drive, it ends with \
                'if it's a folder under drive, it doesn't end with \
                'how shitty
                
                File1.Refresh
        
                Dim fp As String
                fp = File1.path
                If Right(fp, 1) <> "\" Then fp = fp & "\"
                       
                
                If (File1.ListCount < 1 Or fp <> dir) Then
                    socket_send ("0" & DELIMITER)
                    update_log "File list of " & dir & " (no data) sent to " & clientName
                    Exit Sub
                End If
                
                Dim fileList As String
                fileList = ""
                
                For counter = 0 To File1.ListCount - 1
                
                    fh = FreeFile
                    Open dir & File1.List(counter) For Binary As #fh
                    fileList = fileList & File1.List(counter) & " (" & LOF(fh) & "b)" & DELIMITER
                    Close #fh
                    
                Next
                fileList = Left(fileList, Len(fileList) - 1)
                s = Len(fileList) & DELIMITER & fileList
                
            '    MsgBox ("drive: " & Dir1.Path & " s is: " & s)
                socket_send (s)
                update_log "File list of " & dir & " sent to " & clientName
            Else
                socket_send (PX & "You have no permission to view files.")
                update_log clientName & " wants to view files but doesn't have permission."
            End If
        
        Else
            update_log "My state is not CONN but I got PI"
        End If
        
    'client wants to download a file
    Case PL
        If myState = CONN Then
            If (clientPermission And PERMIT_DOWNLOAD) <> 0 Then
              'client requests a file; get file path
                myFilePath = Mid(s, 2)
                
                If (fs.FileExists(myFilePath)) Then
                    'file exists, so send PB and file size
                
                    fh = FreeFile
                    Open myFilePath For Binary As #fh
                    ftpFileSize = LOF(fh)
                    Close #fh
                    
                    If ftpFileSize <= 0 Then
                    
                        socket_send (PX & "File size of " & myFilePath & " is 0 byte.")
                        update_log "File size of " & myFilePath & " is 0 byte."
                     
                    Else
                        pbs.Max = ftpFileSize
                        socket_send (PB & ftpFileSize)
                        update_log clientName & " wants " & myFilePath & " (" & ftpFileSize & " bytes), " & _
                            "PB sent"
                        myState = SEND_FILE
                        speedTimer.Enabled = True
                    
                    End If
                    
                   
                Else
                    'file doesn't exist, send PX
                    'socket_send(PE)
                    socket_send (PX & "Server doesn't have " & myFilePath)
                    myState = CONN
                    update_log clientName & " wants " & myFilePath & " but I don't have it"
                    
                End If
            Else
                socket_send (PX & "You have no permission to download files.")
                update_log clientName & " wants to download files but doesn't have permission."
            End If
        
        Else
            update_log "Not connected."
        End If
        
    Case PA
    'MsgBox "got PA: " & s
        'ack from client, in the middle of ftp
        If (myState = SEND_FILE Or myState = SEND_DIR) Then
            If myState = SEND_DIR And Mid(s, 2) = PA Then
                update_log clientName & " doesn't want " & myFilePath & "; sending next one in line..."
                ftp_file_init
                GoTo Again
            End If
    '    update_log "s: " & s
        
            'it's possible to get "A512A512A1024..." thanks to winsock
            'strip the first A so i got sth like "512A512A1024"
            s = Mid(s, 2)
            Dim ss() As String
            ss() = Split(s, PA) 'u dont split it with DELIMITER...dude...wake up!!
            For counter = LBound(ss) To UBound(ss)
                theirChunkSize = theirChunkSize + Val(ss(counter))
            Next
            
            If (myChunkSize = theirChunkSize) Then
                
                If (currFilePos - 1 = ftpFileSize) Then
                    ftp_file_init
                    update_log "File transfer to client complete"
                    
                    If myState = SEND_FILE Then
                        
                        myState = CONN
                    
                    ElseIf myState = SEND_DIR Then
Again:
                        If filePointer > UBound(allFiles) Then
                            ftp_dir_init
                            myState = CONN
                            'send PE
                            socket_send (PE)
                            update_log "Folder transfer to client complete"
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
                                pbs.Max = ftpFileSize
                                socket_send (PB & myFilePath & DELIMITER & ftpFileSize)
                                update_log clientName & " wants " & myFilePath & " (" & ftpFileSize & " bytes), " & _
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
                Dim bb_len As Long
                finished = get_next_chunk(bb, bb_len)

                socket_send (bb)
                
                'MsgBox "just sent bb: " & bb_len

                If (finished) Then
                
                    update_log "Last packet sent"

                End If
            
            ElseIf (myChunkSize > theirChunkSize) Then
                update_log "My chunk size: " & myChunkSize & " and their chunk " _
                    & "size: " & theirChunkSize & ". Waiting for more Ack..."
            
            Else
                update_log "My chunk size: " & myChunkSize & " and their chunk " _
                    & "size: " & theirChunkSize & ". Forced to terminate FTP because " & _
                    "client claimed to " & vbCrLf & "have received more than we sent."
                
                myState = CONN
                ftp_file_init
                s = "You shouldn't have received more than we sent you."
                socket_send (PX & s)
 
            End If
            
        
            
        Else
            update_log "My state is not SEND_FILE or SEND_DIR but I got PA"
        End If
        
    'client wants to download an entire folder
    Case PZ
   ' MsgBox "msg: " & s
   ' msg is like PZ<0|1><path>DELIMITER<pattern>
        If myState = CONN Then
            If (clientPermission And PERMIT_DOWNLOAD) <> 0 Then
                pos = InStr(1, s, DELIMITER)
                If pos = 0 Then
                    update_log "I got PZ but it doesn't contain pattern: " & s
                    Exit Sub
                End If
                
                'client requests bulk ftp, get its path
                ftpFolderPath = Mid(s, 3, pos - 3)
                
           '     MsgBox "folderpath: " & ftpFolderPath
                
                'get its pattern
                pattern = Mid(s, pos + 1)
                
       'update_log "in pz: ftpfolderpath: " & ftpFolderPath
                'see if i have the dir
                If fs.FolderExists(ftpFolderPath) Then
                    myState = SEND_DIR
                    
       ' update_log "folder exists: " & ftpFolderPath
                    'see if they want recursive or non-recursive
                    'store all files in that folder to allFiles depending on whether or not
                    'it's recursive
                    If Mid(s, 2, 1) = "0" Then
                        'recursive
                        store_all_files_r ftpFolderPath, pattern
                    Else
                        'non-recursive
                        store_all_files ftpFolderPath, pattern
                    End If
                    
                    'take the first one out from allFilse and
           ' MsgBox "count: " & UBound(allFiles)
        'update_log "ubound: " & UBound(allFiles) & "filePointer: " & filePointer
                    filePointer = 0
Again2:
    
                    If filePointer > UBound(allFiles) Then
                        ftp_dir_init
                        myState = CONN
                        'send PE
                        socket_send (PE)
                        update_log "Folder transfer to client complete"
                        Exit Sub
                    End If
                    
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
                            update_log "File size of " & myFilePath & " is 0 byte."
                            GoTo Again2
                         
                        Else
                            pbs.Max = ftpFileSize
                            socket_send (PB & myFilePath & DELIMITER & ftpFileSize)
                            update_log clientName & " wants " & myFilePath & " (" & ftpFileSize & " bytes), " & _
                                "PB sent"
                            myState = SEND_DIR
                            speedTimer.Enabled = True
                        
                        End If
                        
                    Else
                        update_log myFilePath & " doesn't exist; sending next one in line"
                        GoTo Again2
                        
                    End If
                    
                Else
                    'send "E"
                    ws.SendData (PX & "I don't have " & ftpFolderPath)
                    myState = CONN
                    update_log clientName & " wants the entire " & ftpFolderPath & " but I don't have it"
                End If
            Else
                socket_send (PX & "You have no permission to download files.")
                update_log clientName & " wants to download files but doesn't have permission."
            End If
    
        Else
            update_log "My state is not CONN but I got PZ"
        End If
        
    'they want to send server files
    Case PB
    'MsgBox "got a PB, s: " & s
        If myState = CONN Then
            If (clientPermission And PERMIT_UPLOAD) Then
                pos = InStr(2, s, DELIMITER)
                myFilePath = Mid(s, 2, pos - 2)
                
                ftpFileSize = Val(Mid(s, pos + 1))
                'MsgBox "path: " & myFilePath & " size: " & ftpFileSize
                
                If fs.FileExists(myFilePath) Then
                    fs.DeleteFile (myFilePath)
                    update_log myFilePath & " exists, so I deleted it."
                Else
                
                    'fs.GetParentFolderName returns the folder path without ending \
                    pfn = fs.GetParentFolderName(myFilePath)
                    'i need to create the dir if it doesn't exist, for example, if user
                    'wants C:\a\b\c\d\e but I only have C:\a, I need to create C:\a\b
                    'then C:\a\b\c, then C:\a\b\c\d, then C:\a\b\c\d\e
                    If (Mid(PF, Len(pfn)) <> "\") Then pfn = pfn & "\"
                    If (Not fs.FolderExists(pfn)) Then
                        pos = InStr(pfn, "\")
                        Do While pos <> 0
                            pos = InStr(pos + 1, pfn, "\")
                            If (pos = 0) Then Exit Do
                            fd = Mid(pfn, 1, pos)
                            If (Not fs.FolderExists(fd)) Then
                                fs.CreateFolder (fd)
                            End If
                        Loop
                    End If
                
                
                End If
                
                
                
                myState = RECV_FILE
                update_log clientName & " wants to upload " & myFilePath & " (" & ftpFileSize & " b) to me..."
                pbs.Max = ftpFileSize
                socket_send (PA & 0)
                speedTimer.Enabled = True
            Else
                socket_send (PX & "You have no permission to upload files.")
                update_log clientName & " wants to upload files but doesn't have permission."
            End If
        Else
            update_log "My state is not CONN but I got PB: " & s
        End If
        
    Case PQ:
        update_log clientName & " closed connection. Current connection closed."
        'If ws.STATE <> sckClosed Then ws.Close
        reset
        cbss_Click
        
        
    Case PG:
    'PGclientName&clientPassword
        If myState = CONN Then
            pos = InStr(s, DELIMITER)
            clientName = Mid(s, 2, pos - 2)
            clientPassword = Mid(s, pos + 1)
            clientPermission = get_permission(clientName, clientPassword)
            update_log clientName & " has connected me."
            update_log "Permission: " & clientPermission
            If clientPermission < 0 Then
                
                If clientPermission = -1 Then
                    socket_send (PX & "Your account is not found.")
                    update_log "The following client wants to log in but has no account: " & _
                    vbCrLf & "Name: " & clientName & vbCrLf & "Password: " & _
                    clientPassword
                Else
                    socket_send (PX & "Your password is incorrect.")
                    update_log "The following client wants to log in but gave wrong password: " & _
                    vbCrLf & "Name: " & clientName & vbCrLf & "Password: " & _
                    clientPassword
                End If
                DoEvents
                'If ws.STATE <> sckClosed Then ws.Close
                update_log "Current connection closed."
                reset
                cbss_Click
            Else
                'meaning client has an account and its password is correct
                socket_send (PY)
            End If
        Else
            update_log "My state is not CONN but I got PG: " & s
        End If
        
    
        
        
End Select

End Sub

Private Sub ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    update_log ("Error: " & Description)
    reset
    
End Sub
