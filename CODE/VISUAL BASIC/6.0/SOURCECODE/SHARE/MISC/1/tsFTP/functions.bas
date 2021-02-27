Attribute VB_Name = "Module1"
Option Explicit

Function get_socket_state(n As Integer) As String

'ckClosed = 0 sckClosing = 8 sckConnected = 7 sckConnecting =
' sckConnectionPending = 3 sckError = 9 sckHostResolved = 5 sckListening = 2
Select Case n
    Case 0: get_socket_state = "sckClosed"
    Case 1: get_socket_state = "sckOpen"
    Case 2: get_socket_state = "sckListening"
    Case 3: get_socket_state = "sckConnectionPending"
    Case 4: get_socket_state = "sckResolvingHost"
    Case 5: get_socket_state = "sckHostResolved"
    Case 6: get_socket_state = "sckConnecting"
    Case 7: get_socket_state = "sckConnected"
    Case 8: get_socket_state = "sckClosing"
    Case 9: get_socket_state = "sckError"
End Select

End Function

'return true if file name is valid, false otherwise
Public Function is_filename_valid(fn As String) As Boolean

Dim i As Integer
For i = 1 To Len(fn)
    If Mid(fn, i, 1) = "\" Or Mid(fn, i, 1) = "/" Or Mid(fn, i, 1) = ":" Or _
        Mid(fn, i, 1) = "*" Or Mid(fn, i, 1) = "?" Or Mid(fn, i, 1) = "<" Or _
        Mid(fn, i, 1) = ">" Or Mid(fn, i, 1) = "|" Or Mid(fn, i, 1) = """" Then
        is_filename_valid = False
        Exit Function
    End If
Next

is_filename_valid = True
End Function

'store all files recursively (sub dirs) to allFiles
Public Sub store_all_files_r(ByVal start_dir As String, ByVal pattern As String)
Dim dir_names() As String
Dim num_dirs As Integer
Dim i As Integer
Dim fname As String
Dim attr As Integer

'MsgBox "in store all files r"
    ' Protects against such things as
    ' GetAttr("C:\pagefile.sys").
    On Error Resume Next
    
    ' Get the matching files in this directory.
    'however, this wont get hidden or system files...
    fname = Dir(start_dir & "\" & pattern, vbNormal)
    Do While fname <> ""
       'DoEvents
        'If stopList Then Exit Sub
        'lst.AddItem start_dir & "\" & fname
        'MsgBox "fname: " & fname
       ' MsgBox "ubound: " & UBound(allFiles)
        ReDim Preserve allFiles(UBound(allFiles) + 1) As String
        allFiles(UBound(allFiles)) = start_dir & "\" & fname
        fname = Dir()
        'MsgBox "ubound2: " & UBound(allFiles)
    Loop

    ' Get the list of subdirectories.
    fname = Dir(start_dir & "\*.*", vbDirectory)
    Do While fname <> ""
        ' Skip this dir and its parent.
        attr = 0    ' In case there's an error.
        attr = GetAttr(start_dir & "\" & fname)
        If fname <> "." And fname <> ".." And _
            (attr And vbDirectory) <> 0 _
        Then
            num_dirs = num_dirs + 1
            ReDim Preserve dir_names(1 To num_dirs)
            dir_names(num_dirs) = fname
        End If
        fname = Dir()
    Loop
    
    'DoEvents
    'If stopList Then Exit Sub

    ' Search the other directories.
    For i = 1 To num_dirs
        store_all_files_r start_dir & "\" & dir_names(i), pattern
    Next i
End Sub



'store all files to allFiles
Public Sub store_all_files(ByVal start_dir As String, ByVal pattern As String)

Dim dir_names() As String
Dim num_dirs As Integer
Dim i As Integer
Dim fname As String
Dim attr As Integer


    ' Protects against such things as
    ' GetAttr("C:\pagefile.sys").
    On Error Resume Next
    
    ' Get the matching files in this directory.
    fname = Dir(start_dir & pattern, vbNormal)
    Do While fname <> ""
        'DoEvents
        'If stopList Then Exit Sub
        'lst.AddItem start_dir & "\" & fname
        ReDim Preserve allFiles(UBound(allFiles) + 1) As String
        allFiles(UBound(allFiles)) = start_dir & fname
        fname = Dir()
    Loop



End Sub

Public Function get_size(path As String) As Long

Dim fs As New FileSystemObject
If Not fs.FileExists(path) Then
    'MsgBox path & " doesn't exist; you shouldn't invoke get_size()"
    MsgBox path & " doesn't exist; you shouldn't invoke get_size()"
    get_size = 0
    Exit Function
End If

Dim f As File
Set f = fs.GetFile(path)
get_size = f.Size

End Function

Public Sub before_connect_controls()

'disable_all_controls
invisible_all_controls
Form1.mmain.Visible = True
Form1.mexit.Visible = True
Form1.mfolder.Visible = True
Form1.mfile.Visible = True
Form1.mfolder.Enabled = False
Form1.mfile.Enabled = False
Form1.mabout.Visible = True
Form1.mabout2.Visible = True
Form1.mhelp.Visible = True

Form1.Line1.Visible = True
Form1.Line2.Visible = True
Form1.li.Visible = True
Form1.lp.Visible = True
Form1.tbip.Visible = True
Form1.tbp.Visible = True
Form1.cbc.Visible = True
Form1.ls.Visible = True
Form1.ls.Caption = "Click on 'Log in' to log into the remote FTP server"
Form1.cbs.Visible = True

Form1.Height = Form1.cbs.Top + Form1.cbs.Height + UNIT * 3 + Form1.Height - Form1.ScaleHeight

End Sub

Public Sub after_connect_controls()

'disable_all_controls
visible_all_controls
Form1.mfile.Enabled = True
Form1.mfolder.Enabled = True
Form1.li.Visible = False
Form1.lp.Visible = False
Form1.tbip.Visible = False
Form1.tbp.Visible = False
Form1.cbc.Visible = False
Form1.ls.Visible = False
Form1.cbs.Visible = False

Form1.Height = Form1.tbl.Top + Form1.tbl.Height + UNIT * 3 + Form1.Height - Form1.ScaleHeight

End Sub

Public Sub invisible_all_controls()

Dim c As Control
On Error Resume Next
For Each c In Form1.Controls
    c.Visible = False
Next

End Sub

Public Sub visible_all_controls()

Dim c As Control
On Error Resume Next
For Each c In Form1.Controls
    c.Visible = True
Next

End Sub

Public Sub disable_all_controls()

Dim c As Control
On Error Resume Next
For Each c In Form1.Controls
    c.Enabled = False
Next

End Sub

Public Sub enable_all_controls()

Dim c As Control
On Error Resume Next
For Each c In Form1.Controls
    c.Enabled = True
Next

End Sub

Public Sub ftp_init()

    ftp_file_init
    ftp_dir_init

End Sub


Public Sub ftp_file_init()

'bulkFTP = False

ftpFileName = ""
ftpFileSize = 0

myChunkSize = 0
theirChunkSize = 0

myFilePath = ""

Form1.pbs.Min = 0
Form1.pbs.max = 1
Form1.pbs.Value = 0

Form1.speedTimer.Enabled = False
Form1.speedTimer.Interval = 1000
Form1.lspeed = ""
Form1.lbytes = ""


counter = 0

currFilePos = 1 'file pointer always points to position 1 initially

tranOption = TRAN_CANCEL
pattern = ""
defaultFile = ""

'fileFoundOption = UNDECIDED

End Sub


Public Sub ftp_dir_init()
'YOU MUST give an array a size before you can use UBound()
ReDim allFiles(0) As String
pattern = ""
defaultDir = ""
ftpFolderPath = ""
myBasePath = ""
filePointer = 0

End Sub

Public Function in_expansion_list(path As String) As Boolean

Dim counter As Integer
For counter = LBound(expansionList) To UBound(expansionList)
    If expansionList(counter) = path Then
        in_expansion_list = True
        Exit Function
    End If
    
Next counter

in_expansion_list = False

End Function

Public Sub remove_from_expansion_list(path As String)

Dim counter As Integer
For counter = LBound(expansionList) To UBound(expansionList)
    If expansionList(counter) = path Then
        'in_expansion_list = True
        
        Dim counter2 As Integer
        For counter2 = counter To UBound(expansionList) - 1
            expansionList(counter2) = expansionList(counter2 + 1)
        Next
            
        ReDim Preserve expansionList(UBound(expansionList) - 1) As String
            
        Exit Sub
    End If
Next
    
MsgBox (path & " doesn't exist in expansion list...")


End Sub

Public Sub add_to_expansion_list(path As String)

ReDim Preserve expansionList(UBound(expansionList) + 1) As String
expansionList(UBound(expansionList)) = path
        

End Sub

'drivesLen, foldersLen, and filesLen need to be init to -1
'refer to handle(); -1 means it's not init yet, so we need to get
'length from incoming data packet
Public Sub reset_drives()

drivesLen = -1
drivesBuf = ""

End Sub

Public Sub reset_folders()

foldersLen = -1
foldersBuf = ""

End Sub

Public Sub reset_files()

filesLen = -1
filesBuf = ""

'now clear the path
filesRequestPath = ""
numSpaces = 0

End Sub

Public Function bytes_to_string(bb() As Byte) As String

    Dim sAns As String
    Dim iPos As String
    
    sAns = StrConv(bb, vbUnicode)
    iPos = InStr(sAns, Chr(0))
    If iPos > 0 Then sAns = Left(sAns, iPos - 1)
    
    bytes_to_string = sAns
 
End Function
 
Public Function string_to_bytes(s As String)
    
    Dim len_s As Integer
    len_s = Len(s)
    ReDim bb(len_s - 1) As Byte
    Dim i As Integer
    For i = 0 To len_s - 1
        Dim ch As String
        ch = Mid(s, i + 1, 1)
        bb(i) = Asc(ch)
    Next i
    string_to_bytes = bb
    
End Function


