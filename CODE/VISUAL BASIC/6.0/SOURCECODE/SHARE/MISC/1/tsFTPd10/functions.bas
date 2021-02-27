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

Public Function pow(b, e) As Long

Dim i As Long
pow = 1
For i = 1 To e
    pow = pow * b
Next

End Function

'given client's name and password, check AccountFile and get its
'permission bits, return -1 if name doesn't exist, -2 if pw is wrong
Public Function get_permission(name As String, pw As String) As Long

If Not fs.FileExists(AccountFile) Then
    get_permission = -1
Else
    'Dim numEntries As Long
    Dim line As String
    Dim ts As TextStream
    'numEntries = -1
    Set ts = fs.OpenTextFile(AccountFile, ForReading, False)
   
    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        'if the line begins with /, it's a comment
        If Left(line, 1) <> "/" Then
            If line = name Then
               ' MsgBox "line = name"
                If ts.ReadLine <> pw Then
                  '  MsgBox "line <> pw: " & line
                    get_permission = -2
                    ts.Close
                    Exit Function
                Else
                  '  MsgBox "line = pw"
                    'means client gives right info
                    'get permission bits
                    get_permission = 0
                    Dim index As Integer
                    For index = 0 To 5
                        line = ts.ReadLine
                       ' MsgBox "p: " & get_permission & " " & index & " " & pow(2, index)
                        get_permission = get_permission Or (Val(line) * pow(2, index))
                    Next
                    ts.Close
                    Exit Function
                End If
            Else
                'MsgBox "line!= name: " & line
            End If
        End If
    Loop
    
    
    ts.Close
    
    'meaning client's name isn't found in account file
    get_permission = -1
End If

End Function

Public Function get_size(path As String) As Long

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

Public Sub ftp_init()

    ftp_file_init
    ftp_dir_init

End Sub

Public Sub ftp_file_init()

myChunkSize = 0
theirChunkSize = 0
myFilePath = ""
ftpFileName = ""
ftpFilePath = ""
ftpFileSize = 0

Form1.pbs.Min = 0
Form1.pbs.Max = 1
Form1.pbs.Value = 0

Form1.speedTimer.Enabled = False
Form1.speedTimer.Interval = 1000
Form1.lspeed = ""
Form1.lbytes = ""
counter = 0

ftpFinished = False

currFilePos = 1 'file pointer always points to position 1 initially

End Sub

Public Sub ftp_dir_init()


ReDim allFiles(0) As String
ftpFolderPath = "c:\xampp\htdocs\ftp"
filePointer = 0

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

'store all files recursively (sub dirs) to allFiles
Public Sub store_all_files_r(ByVal start_dir As String, ByVal pattern As String)
Dim dir_names() As String
Dim num_dirs As Integer
Dim i As Integer
Dim fname As String
Dim attr As Integer


    ' Protects against such things as
    ' GetAttr("C:\pagefile.sys").
    On Error Resume Next
    
    ' Get the matching files in this directory.
    fname = dir(start_dir & "\" & pattern, vbNormal)
    Do While fname <> ""
        DoEvents
        'If stopList Then Exit Sub
        'lst.AddItem start_dir & "\" & fname
        ReDim Preserve allFiles(UBound(allFiles) + 1) As String
        allFiles(UBound(allFiles)) = start_dir & "\" & fname
        fname = dir()
    Loop

    ' Get the list of subdirectories.
    fname = dir(start_dir & "\*.*", vbDirectory)
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
        fname = dir()
    Loop
    
    DoEvents
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
    fname = dir(start_dir & pattern, vbNormal)
    Do While fname <> ""
        DoEvents
        'If stopList Then Exit Sub
        'lst.AddItem start_dir & "\" & fname
        ReDim Preserve allFiles(UBound(allFiles) + 1) As String
        allFiles(UBound(allFiles)) = start_dir & fname
        fname = dir()
    Loop



End Sub

